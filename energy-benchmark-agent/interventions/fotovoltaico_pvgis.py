"""
Client PVGIS API v5.2 — seriescalc (producibilità oraria PV).

Documentazione ufficiale: https://re.jrc.ec.europa.eu/api/v5_2/
"""

from __future__ import annotations
import time
from typing import List

import requests

_BASE_URL = "https://re.jrc.ec.europa.eu/api/v5_2/seriescalc"
_TIMEOUT = 30  # secondi


def get_producibilita_oraria(
    lat: float,
    lon: float,
    peak_power_kwp: float,
    slope: float,
    aspect: float,
    anno: int = 2023,
    loss: float = 14.0,
    pvtech: str = "crystSi",
    mounting: str = "free",
) -> List[float]:
    """
    Chiama PVGIS seriescalc e restituisce la produzione oraria in kWh per l'anno dato.

    Args:
        lat: latitudine (WGS84)
        lon: longitudine (WGS84)
        peak_power_kwp: potenza di picco dell'array [kWp]
        slope: inclinazione dai orizzontale [°] (0=piano, 90=verticale)
        aspect: orientamento PVGIS [°] (0=sud, -90=est, +90=ovest)
        anno: anno di simulazione (es. 2023)
        loss: perdite totali sistema [%] (default 14%)
        pvtech: tecnologia PV (crystSi | CIS | CdTe | Unknown)
        mounting: tipo montaggio (free | building)

    Returns:
        Lista di 8760 valori [kWh/h] — produzione AC oraria
        (PVGIS restituisce P in W; ogni ora = 1h, quindi W == Wh → /1000 = kWh)

    Raises:
        requests.HTTPError: se la chiamata API fallisce
        ValueError: se la risposta non contiene i dati attesi
    """
    params = {
        "lat": lat,
        "lon": lon,
        "peakpower": peak_power_kwp,
        "loss": loss,
        "angle": slope,
        "aspect": aspect,
        "pvcalculation": 1,          # obbligatorio per ottenere P nel output
        "pvtechchoice": pvtech,
        "mountingplace": mounting,
        "outputformat": "json",
        "startyear": anno,
        "endyear": anno,
        "browser": 0,
    }

    resp = requests.get(_BASE_URL, params=params, timeout=_TIMEOUT)
    resp.raise_for_status()

    data = resp.json()
    hourly = data.get("outputs", {}).get("hourly", [])

    if not hourly:
        raise ValueError(
            f"PVGIS non ha restituito dati orari per lat={lat}, lon={lon}, "
            f"anno={anno}. Risposta: {data}"
        )

    # P è in W; ogni intervallo è 1 ora → P[W] × 1h = Wh → /1000 = kWh
    return [float(h.get("P", 0.0)) / 1000.0 for h in hourly]


def get_producibilita_multi_superficie(
    lat: float,
    lon: float,
    superfici: list,          # List[SuperficieFV] — tipizzato come list per evitare import circolare
    anno: int = 2023,
    loss: float = 14.0,
    pausa_tra_chiamate: float = 0.5,
) -> List[float]:
    """
    Chiama PVGIS una volta per superficie e somma le produzioni orarie.

    Args:
        superfici: lista di SuperficieFV (da interventions.fotovoltaico)
        pausa_tra_chiamate: sleep in secondi tra le chiamate (cortesia verso l'API)

    Returns:
        Lista di 8760 valori [kWh/h] — produzione totale dell'impianto
    """
    totale: List[float] = [0.0] * 8760

    for i, sup in enumerate(superfici):
        if i > 0:
            time.sleep(pausa_tra_chiamate)

        oraria = get_producibilita_oraria(
            lat=lat,
            lon=lon,
            peak_power_kwp=sup.peak_power_kwp,
            slope=sup.slope,
            aspect=sup.azimuth,
            anno=anno,
            loss=loss,
            pvtech=sup.pvtech,
            mounting=sup.mounting,
        )

        n = min(len(oraria), 8760)
        for h in range(n):
            totale[h] += oraria[h]

    return totale
