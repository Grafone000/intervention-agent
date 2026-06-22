"""
Intervento di relamping: sostituzione corpi illuminanti con tecnologia LED.

Per aggiungere un nuovo intervento: copia interventions/_template.py,
rinominalo, implementa calcola(), il registry lo troverà automaticamente.

Struttura output:
  - RelampingRiga: calcoli riga per riga (foglio "Calcoli")
  - CBCalcolo:     certificati bianchi (foglio "CB")
  - VanScenario:   indicatori economici per ognuno dei 4 scenari VAN
  - RelampingResult(InterventionResult): oggetto completo
"""

from __future__ import annotations
import math
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field

from core.energy_model import EnergyModel
from interventions.base import InterventionResult
from interventions.relamping_map import converti_lampada

NOME_INTERVENTO = "relamping"

# Parametri fissi dell'intervento
_VITA_ANNI = 8
_DURATA_INCENTIVI_ANNI = 5
_VALORE_CB_EURO_TEP = 250.0
_DISCOUNT_RATE = 0.06
_FATTORE_TEP = 0.000187
_FATTORE_CO2 = 294.784 / 1000   # kg/kWh
_PREZZO_KWH = 0.21               # da FATTORI CONVERSIONE C24


# ---------------------------------------------------------------------------
# Modelli dati output
# ---------------------------------------------------------------------------

class RelampingRiga(BaseModel):
    """Risultato riga per riga — corrisponde al foglio 'Calcoli'."""
    pod: str
    zona: str
    tipologia_originale: str
    categoria_led: str
    n_utenze: int
    p_lampada_ante_kw: float
    p_tot_ante_kw: float
    p_lampada_post_kw: float
    costo_lampada_euro: float
    p_tot_post_kw: float
    costo_totale_euro: float
    h_equivalenti: float
    fatt_contemporaneita: float
    consumo_ante_kwh: float
    consumo_post_kwh: float
    spesa_ante_euro: float
    spesa_post_euro: float
    co2_ante_kg: float
    co2_post_kg: float
    risparmio_ee_kwh: float
    risparmio_euro: float
    risparmio_co2_kg: float
    confidenza: str
    note: List[str] = []


class CBCalcolo(BaseModel):
    """Calcolo certificati bianchi — corrisponde al foglio 'CB'."""
    p_tot_ante_kw: float = Field(description="Somma P tot ante di tutte le righe")
    e_post_kwh: float = Field(description="Consumo totale post")
    p_tot_post_kw: float = Field(description="Somma P tot post")
    h_post: float = Field(description="Ore funzionamento post = E_post / P_tot_post")
    tep_risparmiati: float = Field(description="(P_ant×h_post − E_post) × 0.000187")
    tep_floor: int = Field(description="tep_risparmiati approssimati per difetto")
    incentivo_annuo: float = Field(description="tep_floor × 250 €/tep")


class VanScenario(BaseModel):
    """Indicatori economici per uno dei 4 scenari VAN."""
    nome: str
    investimento: float
    risparmio_annuo_ee_kwh: float
    risparmio_annuo_euro: float
    discount_rate: float
    include_incentivi: bool
    incentivo_annuo: float
    risparmi_per_anno: List[float] = Field(description="Risparmio (€) per ciascuno degli 8 anni")
    fattori_attualizzazione: List[float]
    flussi_cassa: List[float] = Field(description="FC cumulato per ciascun anno")
    van: float
    tr: float
    dpp: Optional[float]
    tir: Optional[float]
    indice_profitto: float


class RelampingResult(InterventionResult):
    """Risultato completo dell'intervento di relamping."""
    righe: List[RelampingRiga]
    cb: CBCalcolo
    van_attualizzato_con_incentivi: VanScenario
    van_attualizzato_senza_incentivi: VanScenario
    van_semplice_con_incentivi: VanScenario
    van_semplice_senza_incentivi: VanScenario
    prezzo_kwh: float = _PREZZO_KWH
    fattore_co2: float = _FATTORE_CO2


# ---------------------------------------------------------------------------
# Funzioni di calcolo
# ---------------------------------------------------------------------------

def _calcola_van_scenario(
    investimento: float,
    risparmio_annuo_euro: float,
    risparmio_annuo_ee_kwh: float,
    incentivo_annuo: float,
    discount_rate: float,
    include_incentivi: bool,
    nome: str,
) -> VanScenario:
    """Calcola flussi di cassa, VAN, TR, DPP, TIR, IP per uno scenario."""

    risparmi: list[float] = []
    for anno in range(1, _VITA_ANNI + 1):
        r = risparmio_annuo_euro
        if include_incentivi and anno <= _DURATA_INCENTIVI_ANNI:
            r += incentivo_annuo
        risparmi.append(r)

    # Fattori di attualizzazione f_att_n = 1/(1+r)^n
    f_att = [1.0 / (1 + discount_rate) ** n for n in range(1, _VITA_ANNI + 1)]

    # Flussi di cassa cumulati
    # FC_n = -Inv + Σ_{k=1}^{n}(risp_k × f_att_k)
    flussi: list[float] = []
    cumulo_attualizzato = 0.0
    for n in range(_VITA_ANNI):
        cumulo_attualizzato += risparmi[n] * f_att[n]
        fc_n = -investimento + cumulo_attualizzato
        flussi.append(round(fc_n, 4))

    van = flussi[-1]

    # Tempo di ritorno semplice
    if risparmio_annuo_euro <= 0:
        tr = float("inf")
    elif include_incentivi:
        tr = (investimento - incentivo_annuo * _DURATA_INCENTIVI_ANNI) / risparmio_annuo_euro
    else:
        tr = investimento / risparmio_annuo_euro

    # Tempo di ritorno attualizzato (DPP)
    dpp: Optional[float] = None
    for n in range(len(flussi)):
        if flussi[n] >= 0:
            if n == 0:
                dpp = 0.0
            else:
                fc_prev = flussi[n - 1]
                fc_curr = flussi[n]
                dpp = (n - 1) + abs(fc_prev) / (abs(fc_curr) + abs(fc_prev))
            break

    # TIR — solo per scenari attualizzati (discount_rate > 0)
    tir: Optional[float] = None
    if discount_rate > 0:
        # Cash flow incrementali: CF_0 = -Inv, CF_n = risparmio_n
        cf_incrementali = [-investimento] + risparmi
        tir = _calcola_tir(cf_incrementali)

    ip = van / investimento if investimento > 0 else 0.0

    return VanScenario(
        nome=nome,
        investimento=investimento,
        risparmio_annuo_ee_kwh=risparmio_annuo_ee_kwh,
        risparmio_annuo_euro=risparmio_annuo_euro,
        discount_rate=discount_rate,
        include_incentivi=include_incentivi,
        incentivo_annuo=incentivo_annuo,
        risparmi_per_anno=risparmi,
        fattori_attualizzazione=f_att,
        flussi_cassa=flussi,
        van=round(van, 2),
        tr=round(tr, 2),
        dpp=round(dpp, 2) if dpp is not None else None,
        tir=round(tir * 100, 2) if tir is not None else None,
        indice_profitto=round(ip, 4),
    )


def _calcola_tir(cash_flows: list[float], tol: float = 1e-7, max_iter: int = 2000) -> Optional[float]:
    """Calcola il TIR con metodo di bisezione."""
    def npv(rate: float) -> float:
        return sum(cf / (1 + rate) ** t for t, cf in enumerate(cash_flows))

    try:
        low, high = -0.9999, 50.0
        npv_low = npv(low)
        npv_high = npv(high)
        if npv_low * npv_high > 0:
            return None
        for _ in range(max_iter):
            mid = (low + high) / 2.0
            npv_mid = npv(mid)
            if abs(npv_mid) < tol or (high - low) / 2 < tol:
                return mid
            if npv_low * npv_mid < 0:
                high = mid
                npv_high = npv_mid
            else:
                low = mid
                npv_low = npv_mid
        return (low + high) / 2.0
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Funzione principale
# ---------------------------------------------------------------------------

def calcola(model: EnergyModel, parametri: Dict[str, Any]) -> RelampingResult:
    """
    Calcola il benchmark dell'intervento di relamping.

    Args:
        model: modello energetico (da core/parser.py)
        parametri: dizionario opzionale, chiavi supportate:
            - prezzo_kwh (float): sovrascrive il default 0.21 €/kWh
            - valore_cb (float): sovrascrive il default 250 €/tep

    Returns:
        RelampingResult con calcoli riga per riga, CB e indicatori economici
    """
    prezzo_kwh = float(parametri.get("prezzo_kwh", _PREZZO_KWH))
    valore_cb = float(parametri.get("valore_cb", _VALORE_CB_EURO_TEP))

    candidati = model.utenze_relamping()
    if not candidati:
        raise ValueError(
            "Nessuna utenza candidata al relamping trovata nel modello. "
            "Verifica che ci siano utenze con uso_energetico='Illuminazione' "
            "e tipologia senza 'LED'."
        )

    righe: list[RelampingRiga] = []

    for u in candidati:
        conv = converti_lampada(u.tipologia, u.potenza_unitaria_kw)

        p_tot_ante = u.potenza_unitaria_kw * u.n_utenze
        p_tot_post = conv.p_post_kw * u.n_utenze
        costo_totale = conv.costo_lampada_euro * u.n_utenze

        consumo_ante = p_tot_ante * u.fatt_contemporaneita * u.h_equivalenti
        consumo_post = p_tot_post * u.fatt_contemporaneita * u.h_equivalenti

        spesa_ante = consumo_ante * prezzo_kwh
        spesa_post = consumo_post * prezzo_kwh

        co2_ante = consumo_ante * _FATTORE_CO2
        co2_post = consumo_post * _FATTORE_CO2

        righe.append(RelampingRiga(
            pod=u.pod,
            zona=u.edificio_zona,
            tipologia_originale=u.tipologia,
            categoria_led=conv.categoria,
            n_utenze=u.n_utenze,
            p_lampada_ante_kw=u.potenza_unitaria_kw,
            p_tot_ante_kw=round(p_tot_ante, 6),
            p_lampada_post_kw=round(conv.p_post_kw, 6),
            costo_lampada_euro=conv.costo_lampada_euro,
            p_tot_post_kw=round(p_tot_post, 6),
            costo_totale_euro=round(costo_totale, 2),
            h_equivalenti=u.h_equivalenti,
            fatt_contemporaneita=u.fatt_contemporaneita,
            consumo_ante_kwh=round(consumo_ante, 4),
            consumo_post_kwh=round(consumo_post, 4),
            spesa_ante_euro=round(spesa_ante, 2),
            spesa_post_euro=round(spesa_post, 2),
            co2_ante_kg=round(co2_ante, 4),
            co2_post_kg=round(co2_post, 4),
            risparmio_ee_kwh=round(consumo_ante - consumo_post, 4),
            risparmio_euro=round(spesa_ante - spesa_post, 2),
            risparmio_co2_kg=round(co2_ante - co2_post, 4),
            confidenza=conv.confidenza,
            note=conv.note + [n for n in u.note_parser if n],
        ))

    # --- Aggregati ---
    p_tot_ante_totale = sum(r.p_tot_ante_kw for r in righe)
    e_post_totale = sum(r.consumo_post_kwh for r in righe)
    p_tot_post_totale = sum(r.p_tot_post_kw for r in righe)
    investimento_totale = sum(r.costo_totale_euro for r in righe)
    risparmio_ee_totale = sum(r.risparmio_ee_kwh for r in righe)
    risparmio_euro_totale = sum(r.risparmio_euro for r in righe)

    # --- CB (Certificati Bianchi) ---
    h_post = e_post_totale / p_tot_post_totale if p_tot_post_totale > 0 else 0.0
    tep_risparmiati = (p_tot_ante_totale * h_post - e_post_totale) * _FATTORE_TEP
    tep_floor = math.floor(tep_risparmiati)
    incentivo_annuo = tep_floor * valore_cb

    cb = CBCalcolo(
        p_tot_ante_kw=round(p_tot_ante_totale, 4),
        e_post_kwh=round(e_post_totale, 4),
        p_tot_post_kw=round(p_tot_post_totale, 4),
        h_post=round(h_post, 2),
        tep_risparmiati=round(tep_risparmiati, 6),
        tep_floor=tep_floor,
        incentivo_annuo=round(incentivo_annuo, 2),
    )

    # --- 4 scenari VAN ---
    scenari_config = [
        ("VAN Attualizzato con incentivi",    _DISCOUNT_RATE, True),
        ("VAN Attualizzato senza incentivi",  _DISCOUNT_RATE, False),
        ("VAN Semplice con incentivi",        0.0,            True),
        ("VAN Semplice senza incentivi",      0.0,            False),
    ]
    van_results: list[VanScenario] = []
    for nome, dr, inc in scenari_config:
        van_results.append(_calcola_van_scenario(
            investimento=investimento_totale,
            risparmio_annuo_euro=risparmio_euro_totale,
            risparmio_annuo_ee_kwh=risparmio_ee_totale,
            incentivo_annuo=incentivo_annuo,
            discount_rate=dr,
            include_incentivi=inc,
            nome=nome,
        ))

    note_generali = []
    righe_bassa_conf = [r for r in righe if r.confidenza in ("Bassa", "Da verificare")]
    if righe_bassa_conf:
        note_generali.append(
            f"{len(righe_bassa_conf)} righe con confidenza Bassa/Da verificare — "
            "verificare manualmente le potenze post stimate"
        )

    return RelampingResult(
        nome_intervento=NOME_INTERVENTO,
        risparmio_annuo_kwh=round(risparmio_ee_totale, 2),
        costo_stimato_euro=round(investimento_totale, 2),
        payback_anni=round(van_results[1].tr, 2),  # TR attualizzato senza incentivi
        dati_grafico={
            "risparmio_ee_kwh": risparmio_ee_totale,
            "risparmio_euro": risparmio_euro_totale,
            "co2_risparmiata_kg": sum(r.risparmio_co2_kg for r in righe),
            "investimento": investimento_totale,
            "van_con_incentivi": van_results[0].van,
            "van_senza_incentivi": van_results[1].van,
        },
        note=note_generali,
        righe=righe,
        cb=cb,
        van_attualizzato_con_incentivi=van_results[0],
        van_attualizzato_senza_incentivi=van_results[1],
        van_semplice_con_incentivi=van_results[2],
        van_semplice_senza_incentivi=van_results[3],
        prezzo_kwh=prezzo_kwh,
        fattore_co2=_FATTORE_CO2,
    )
