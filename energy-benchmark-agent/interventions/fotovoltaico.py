"""
Intervento fotovoltaico: installazione impianto FV con calcolo producibilità via PVGIS.

Input richiesti (in parametri):
    lat (float)              : latitudine sito
    lon (float)              : longitudine sito
    superfici (list[dict])   : lista superfici FV
                               ciascuna: {id, slope, azimuth, n_pannelli, potenza_wp,
                                          pvtech='crystSi', mounting='free'}
    percorso_consumi (str)   : percorso file consumi quart'orari (CSV o Excel)
    anno_pvgis (int)         : anno simulazione PVGIS, default 2023
    loss (float)             : perdite sistema [%], default 14.0
    prezzo_kwh (float)       : prezzo acquisto energia [€/kWh], default 0.21
    pun_euro_kwh (float)     : prezzo cessione in rete [€/kWh], default 0.12
    valore_cb (float)        : valore CB [€/tep], default 250.0
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field

from core.energy_model import EnergyModel
from core.consumi_parser import parse_consumi_quart_orari, parse_consumi_excel_multi_pod
from interventions.base import InterventionResult
from interventions.economics import (
    CBCalcolo, VanScenario, calcola_cb, calcola_van_scenario,
    _FATTORE_TEP, _DISCOUNT_RATE, _DURATA_INCENTIVI_ANNI,
)
from interventions.fotovoltaico_pvgis import get_producibilita_oraria
from interventions.fotovoltaico_prezzi import costo_impianto

NOME_INTERVENTO = "fotovoltaico"

_PREZZO_KWH_DEFAULT = 0.21
_PUN_DEFAULT = 0.12
_FATTORE_CO2 = 294.784 / 1000   # kg/kWh


# ---------------------------------------------------------------------------
# Modelli dati input
# ---------------------------------------------------------------------------

class SuperficieFV(BaseModel):
    """Superficie fotovoltaica (falda, pensilina, ecc.)."""
    id: str = Field(description="Identificativo superficie (es. 'Tetto Sud')")
    slope: float = Field(description="Inclinazione da orizzontale [°]: 0=piano, 90=verticale")
    azimuth: float = Field(description="Orientamento PVGIS [°]: 0=sud, -90=est, +90=ovest")
    n_pannelli: int = Field(description="Numero di pannelli")
    potenza_pannello_wp: float = Field(description="Potenza singolo pannello [Wp]")
    pvtech: str = Field(default="crystSi", description="Tecnologia: crystSi|CIS|CdTe|Unknown")
    mounting: str = Field(default="free", description="Tipo montaggio: free|building")

    @property
    def peak_power_kwp(self) -> float:
        return self.n_pannelli * self.potenza_pannello_wp / 1000.0


# ---------------------------------------------------------------------------
# Modelli dati output
# ---------------------------------------------------------------------------

class FVRiga(BaseModel):
    """Producibilità per singola superficie."""
    superficie_id: str
    slope: float
    azimuth: float
    n_pannelli: int
    potenza_pannello_wp: float
    peak_power_kwp: float
    e_prodotta_kwh: float
    ore_equivalenti: float


class FVResult(InterventionResult):
    """Risultato completo dell'intervento fotovoltaico."""
    superfici_input: List[SuperficieFV]
    righe: List[FVRiga]

    # Dati impianto
    potenza_totale_kwp: float
    costo_per_kwp: float

    # Producibilità
    e_prodotta_kwh: float
    ore_equivalenti_impianto: float

    # Autoconsumo
    e_autoconsumata_kwh: float
    e_immessa_kwh: float
    e_prelevata_kwh: float
    quota_autoconsumo: float

    # Economico annuo
    risparmio_acquisto_euro: float
    ricavo_immissione_euro: float
    risparmio_totale_euro: float
    co2_evitata_kg: float

    # Certificati Bianchi
    cb: CBCalcolo

    # Benchmark VAN
    van_attualizzato_con_incentivi: VanScenario
    van_attualizzato_senza_incentivi: VanScenario
    van_semplice_con_incentivi: VanScenario
    van_semplice_senza_incentivi: VanScenario

    # Parametri usati
    prezzo_kwh: float
    pun_euro_kwh: float
    anno_pvgis: int


# ---------------------------------------------------------------------------
# Parsing tabella superfici
# ---------------------------------------------------------------------------

def _parse_superfici(superfici_raw: list[dict]) -> list[SuperficieFV]:
    """Converte lista di dict (da parametri) in lista di SuperficieFV."""
    result = []
    for d in superfici_raw:
        result.append(SuperficieFV(
            id=str(d.get("id", f"Superficie_{len(result)+1}")),
            slope=float(d["slope"]),
            azimuth=float(d["azimuth"]),
            n_pannelli=int(d["n_pannelli"]),
            potenza_pannello_wp=float(d["potenza_wp"]),
            pvtech=str(d.get("pvtech", "crystSi")),
            mounting=str(d.get("mounting", "free")),
        ))
    return result


def _parse_superfici_file(filepath: str) -> list[SuperficieFV]:
    """Legge la tabella superfici da CSV o Excel."""
    from pathlib import Path
    import csv

    fp = Path(filepath)
    if not fp.exists():
        raise FileNotFoundError(f"File superfici non trovato: {filepath}")

    ext = fp.suffix.lower()
    rows: list[dict] = []

    if ext in (".xlsx", ".xls"):
        import openpyxl
        wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            raise ValueError("File superfici vuoto")
        header = [str(c).strip().lower() if c else f"col{i}" for i, c in enumerate(all_rows[0])]
        for row in all_rows[1:]:
            if row and any(v is not None for v in row):
                rows.append(dict(zip(header, row)))
    else:
        with open(fp, newline="", encoding="utf-8-sig") as f:
            sample = f.read(2048)
            f.seek(0)
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            reader = csv.DictReader(f, dialect=dialect)
            rows = [r for r in reader]

    superfici = []
    for i, row in enumerate(rows):
        # normalizza chiavi
        norm = {k.strip().lower(): v for k, v in row.items() if v not in (None, "")}

        def get(*keys):
            for k in keys:
                if k in norm:
                    return norm[k]
            return None

        id_val = get("id", "nome", "superficie") or f"Superficie_{i+1}"
        slope = get("slope", "tilt", "inclinazione")
        azimuth = get("azimuth", "aspect", "orientamento")
        n = get("n_pannelli", "n pannelli", "pannelli", "n")
        pwr = get("potenza_wp", "potenza wp", "wp", "potenza_pannello_wp", "potenza")
        pvtech = get("pvtech", "tecnologia") or "crystSi"
        mounting = get("mounting", "montaggio") or "free"

        if any(v is None for v in [slope, azimuth, n, pwr]):
            raise ValueError(
                f"Riga {i+2} del file superfici: campi obbligatori mancanti "
                f"(slope, azimuth, n_pannelli, potenza_wp). Trovato: {dict(norm)}"
            )

        superfici.append(SuperficieFV(
            id=str(id_val),
            slope=float(slope),
            azimuth=float(azimuth),
            n_pannelli=int(float(n)),
            potenza_pannello_wp=float(pwr),
            pvtech=str(pvtech),
            mounting=str(mounting),
        ))

    return superfici


# ---------------------------------------------------------------------------
# Calcolo autoconsumo
# ---------------------------------------------------------------------------

def _calcola_autoconsumo(
    produzione: list[float],
    consumo: list[float],
) -> tuple[float, float, float]:
    """
    Calcola autoconsumo, immissione e prelievo annui.

    Returns:
        (e_autoconsumata_kwh, e_immessa_kwh, e_prelevata_kwh)
    """
    auto = imm = prel = 0.0
    for p, c in zip(produzione, consumo):
        auto += min(p, c)
        imm  += max(0.0, p - c)
        prel += max(0.0, c - p)
    return auto, imm, prel


# ---------------------------------------------------------------------------
# Funzione principale
# ---------------------------------------------------------------------------

def calcola(model: EnergyModel, parametri: Dict[str, Any]) -> FVResult:
    """
    Calcola il benchmark dell'intervento fotovoltaico.

    Args:
        model: modello energetico (usato per regione e metadati)
        parametri: dizionario con chiavi:
            lat (float)                 : latitudine sito [obbligatorio]
            lon (float)                 : longitudine sito [obbligatorio]
            superfici (list[dict])      : tabella superfici [obbligatorio se non percorso_superfici]
            percorso_superfici (str)    : percorso CSV/Excel superfici [alternativo a superfici]
            percorso_consumi (str)      : percorso file consumi quart'orari [obbligatorio]
            anno_pvgis (int)            : anno simulazione PVGIS (default 2023)
            loss (float)                : perdite sistema % (default 14.0)
            prezzo_kwh (float)          : prezzo acquisto ee (default 0.21)
            pun_euro_kwh (float)        : prezzo cessione in rete (default 0.12)
            valore_cb (float)           : valore CB €/tep (default 250.0)

    Returns:
        FVResult con tutti i calcoli energetici ed economici
    """
    # --- Parametri ---
    lat = float(parametri["lat"])
    lon = float(parametri["lon"])
    anno = int(parametri.get("anno_pvgis", 2023))
    loss = float(parametri.get("loss", 14.0))
    prezzo_kwh = float(parametri.get("prezzo_kwh", _PREZZO_KWH_DEFAULT))
    pun = float(parametri.get("pun_euro_kwh", _PUN_DEFAULT))
    valore_cb = float(parametri.get("valore_cb", 250.0))

    # --- Superfici ---
    if "superfici" in parametri:
        superfici = _parse_superfici(parametri["superfici"])
    elif "percorso_superfici" in parametri:
        superfici = _parse_superfici_file(str(parametri["percorso_superfici"]))
    else:
        raise ValueError("parametri deve contenere 'superfici' (list) o 'percorso_superfici' (path)")

    if not superfici:
        raise ValueError("Nessuna superficie FV definita")

    # --- Consumi ---
    if "percorso_consumi" not in parametri:
        raise ValueError("parametri deve contenere 'percorso_consumi'")
    percorso_consumi = str(parametri["percorso_consumi"])
    # Auto-detect formato: Excel multi-POD (foglio "Input POD orario ATTIVA") o CSV quart'orario
    if percorso_consumi.lower().endswith((".xlsx", ".xls")):
        try:
            import openpyxl
            wb_check = openpyxl.load_workbook(percorso_consumi, data_only=True, read_only=True)
            foglio_multi = "Input POD orario ATTIVA"
            if foglio_multi in wb_check.sheetnames:
                col_prima_pod = int(parametri.get("col_prima_pod", 130))
                consumo_orario = parse_consumi_excel_multi_pod(
                    percorso_consumi, nome_foglio=foglio_multi, col_prima_pod=col_prima_pod
                )
            else:
                consumo_orario = parse_consumi_quart_orari(percorso_consumi)
        except Exception:
            consumo_orario = parse_consumi_quart_orari(percorso_consumi)
    else:
        consumo_orario = parse_consumi_quart_orari(percorso_consumi)

    # --- PVGIS: producibilità oraria per superficie e totale ---
    import time

    produzione_totale: list[float] = [0.0] * 8760
    righe: list[FVRiga] = []

    for i, sup in enumerate(superfici):
        if i > 0:
            time.sleep(0.5)
        prod_sup = get_producibilita_oraria(
            lat=lat, lon=lon,
            peak_power_kwp=sup.peak_power_kwp,
            slope=sup.slope, aspect=sup.azimuth,
            anno=anno, loss=loss,
            pvtech=sup.pvtech, mounting=sup.mounting,
        )
        n = min(len(prod_sup), 8760)
        for h in range(n):
            produzione_totale[h] += prod_sup[h]
        e_sup = sum(prod_sup[:n])
        h_eq = e_sup / sup.peak_power_kwp if sup.peak_power_kwp > 0 else 0.0
        righe.append(FVRiga(
            superficie_id=sup.id,
            slope=sup.slope,
            azimuth=sup.azimuth,
            n_pannelli=sup.n_pannelli,
            potenza_pannello_wp=sup.potenza_pannello_wp,
            peak_power_kwp=round(sup.peak_power_kwp, 3),
            e_prodotta_kwh=round(e_sup, 1),
            ore_equivalenti=round(h_eq, 1),
        ))

    # --- Aggregati impianto ---
    potenza_totale = sum(s.peak_power_kwp for s in superfici)
    e_prodotta = sum(produzione_totale)
    ore_eq_impianto = e_prodotta / potenza_totale if potenza_totale > 0 else 0.0

    # --- Autoconsumo ---
    e_auto, e_imm, e_prel = _calcola_autoconsumo(produzione_totale, consumo_orario)
    quota_auto = e_auto / e_prodotta if e_prodotta > 0 else 0.0

    # --- Economico annuo ---
    risp_acquisto = e_auto * prezzo_kwh
    ricavo_imm = e_imm * pun
    risp_totale = risp_acquisto + ricavo_imm
    co2_evitata = e_prodotta * _FATTORE_CO2

    # --- Investimento ---
    costo_totale = costo_impianto(potenza_totale, model.regione)
    costo_unitario = costo_totale / potenza_totale if potenza_totale > 0 else 0.0

    # --- Certificati Bianchi ---
    tep_risparmiati = e_prodotta * _FATTORE_TEP
    cb = calcola_cb(tep_risparmiati, valore_cb)

    # --- 4 scenari VAN ---
    scenari_config = [
        ("VAN Attualizzato con incentivi",   _DISCOUNT_RATE, True),
        ("VAN Attualizzato senza incentivi", _DISCOUNT_RATE, False),
        ("VAN Semplice con incentivi",       0.0,            True),
        ("VAN Semplice senza incentivi",     0.0,            False),
    ]
    van_results: list[VanScenario] = []
    for nome, dr, inc in scenari_config:
        van_results.append(calcola_van_scenario(
            investimento=costo_totale,
            risparmio_annuo_euro=risp_totale,
            risparmio_annuo_ee_kwh=e_auto,
            incentivo_annuo=cb.incentivo_annuo,
            discount_rate=dr,
            include_incentivi=inc,
            nome=nome,
        ))

    return FVResult(
        nome_intervento=NOME_INTERVENTO,
        risparmio_annuo_kwh=round(e_auto, 1),
        costo_stimato_euro=round(costo_totale, 2),
        payback_anni=round(van_results[1].tr, 2),
        dati_grafico={
            "e_prodotta_kwh": e_prodotta,
            "e_autoconsumata_kwh": e_auto,
            "e_immessa_kwh": e_imm,
            "quota_autoconsumo": quota_auto,
            "risparmio_totale_euro": risp_totale,
            "co2_evitata_kg": co2_evitata,
            "investimento": costo_totale,
            "van_con_incentivi": van_results[0].van,
            "van_senza_incentivi": van_results[1].van,
        },
        note=[],
        superfici_input=superfici,
        righe=righe,
        potenza_totale_kwp=round(potenza_totale, 3),
        costo_per_kwp=round(costo_unitario, 2),
        e_prodotta_kwh=round(e_prodotta, 1),
        ore_equivalenti_impianto=round(ore_eq_impianto, 1),
        e_autoconsumata_kwh=round(e_auto, 1),
        e_immessa_kwh=round(e_imm, 1),
        e_prelevata_kwh=round(e_prel, 1),
        quota_autoconsumo=round(quota_auto, 4),
        risparmio_acquisto_euro=round(risp_acquisto, 2),
        ricavo_immissione_euro=round(ricavo_imm, 2),
        risparmio_totale_euro=round(risp_totale, 2),
        co2_evitata_kg=round(co2_evitata, 1),
        cb=cb,
        van_attualizzato_con_incentivi=van_results[0],
        van_attualizzato_senza_incentivi=van_results[1],
        van_semplice_con_incentivi=van_results[2],
        van_semplice_senza_incentivi=van_results[3],
        prezzo_kwh=prezzo_kwh,
        pun_euro_kwh=pun,
        anno_pvgis=anno,
    )
