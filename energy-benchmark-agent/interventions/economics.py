"""
Modelli e funzioni economiche condivise tra gli interventi.

Esportato da qui: VanScenario, CBCalcolo, calcola_van_scenario, calcola_tir
"""

from __future__ import annotations
from typing import List, Optional

from pydantic import BaseModel, Field

_VITA_ANNI = 8
_DURATA_INCENTIVI_ANNI = 5
_VALORE_CB_EURO_TEP = 250.0
_DISCOUNT_RATE = 0.06
_FATTORE_TEP = 0.000187


class CBCalcolo(BaseModel):
    """Calcolo certificati bianchi generico."""
    tep_risparmiati: float
    tep_arrotondati: int = Field(description="ROUND(tep_risparmiati, 0)")
    incentivo_annuo: float = Field(description="tep_arrotondati × 250 €/tep")


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


def calcola_tir(cash_flows: list[float], tol: float = 1e-7, max_iter: int = 2000) -> Optional[float]:
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


def calcola_van_scenario(
    investimento: float,
    risparmio_annuo_euro: float,
    risparmio_annuo_ee_kwh: float,
    incentivo_annuo: float,
    discount_rate: float,
    include_incentivi: bool,
    nome: str,
    vita_anni: int = _VITA_ANNI,
    durata_incentivi_anni: int = _DURATA_INCENTIVI_ANNI,
) -> VanScenario:
    """Calcola flussi di cassa, VAN, TR, DPP, TIR, IP per uno scenario."""

    risparmi: list[float] = []
    for anno in range(1, vita_anni + 1):
        r = risparmio_annuo_euro
        if include_incentivi and anno <= durata_incentivi_anni:
            r += incentivo_annuo
        risparmi.append(r)

    f_att = [1.0 / (1 + discount_rate) ** n for n in range(1, vita_anni + 1)]

    flussi: list[float] = []
    cumulo_attualizzato = 0.0
    for n in range(vita_anni):
        cumulo_attualizzato += risparmi[n] * f_att[n]
        fc_n = -investimento + cumulo_attualizzato
        flussi.append(round(fc_n, 4))

    van = flussi[-1]

    if risparmio_annuo_euro <= 0:
        tr = float("inf")
    elif include_incentivi:
        tr = (investimento - incentivo_annuo * durata_incentivi_anni) / risparmio_annuo_euro
    else:
        tr = investimento / risparmio_annuo_euro

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

    tir: Optional[float] = None
    if discount_rate > 0:
        cf_incrementali = [-investimento] + risparmi
        tir = calcola_tir(cf_incrementali)

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


def calcola_cb(tep_risparmiati: float, valore_cb: float = _VALORE_CB_EURO_TEP) -> CBCalcolo:
    tep_arr = round(tep_risparmiati)
    return CBCalcolo(
        tep_risparmiati=round(tep_risparmiati, 6),
        tep_arrotondati=tep_arr,
        incentivo_annuo=round(tep_arr * valore_cb, 2),
    )
