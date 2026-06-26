"""
Prezziario impianti fotovoltaici per taglia e regione.

Fonti: Sintesi_elenco_prezzi.md (prezziario regionale) + medie di mercato 2024-2025.
"""

from __future__ import annotations

# Fasce mercato 2024-2025 (€/kWp) — default quando la regione non è nella tabella
_FASCE_MERCATO: list[tuple[float, float]] = [
    # (soglia_kwp, euro_per_kwp)
    (6.0,    1500.0),
    (20.0,   1250.0),
    (100.0,  1050.0),
    (float("inf"), 800.0),
]

# Prezziario regionale — da Sintesi_elenco_prezzi.md
# Struttura: {regione_lower: [(soglia_kwp, euro_per_kwp), ...]}
_PREZZI_REGIONALI: dict[str, list[tuple[float, float]]] = {
    "lombardia": [
        (6.0,         3580.56),
        (20.0,        2800.00),
        (50.0,        2500.00),
        (float("inf"), 2288.91),
    ],
    "marche": [
        (5.0,         3132.28),
        (20.0,        2200.00),
        (100.0,       1600.00),
        (500.0,       1350.00),
        (1000.0,      1256.41),
        (float("inf"), 1256.41),
    ],
    "sardegna": [
        (6.0,         2284.46),
        (20.0,        2100.00),
        (50.0,        1900.00),
        (float("inf"), 1707.90),
    ],
}


def costo_per_kwp(potenza_kwp: float, regione: str | None = None) -> float:
    """
    Restituisce il costo unitario [€/kWp] in funzione della taglia e della regione.

    Args:
        potenza_kwp: potenza di picco dell'impianto [kWp]
        regione: regione del sito (case-insensitive, es. "Sardegna"). None = default mercato.

    Returns:
        Costo unitario [€/kWp]
    """
    tabella = _FASCE_MERCATO
    if regione:
        key = regione.lower().strip()
        if key in _PREZZI_REGIONALI:
            tabella = _PREZZI_REGIONALI[key]

    for soglia, prezzo in tabella:
        if potenza_kwp <= soglia:
            return prezzo
    return tabella[-1][1]


def costo_impianto(potenza_kwp: float, regione: str | None = None) -> float:
    """
    Costo totale dell'impianto [€] = potenza_kwp × costo_per_kwp.
    """
    return round(potenza_kwp * costo_per_kwp(potenza_kwp, regione), 2)
