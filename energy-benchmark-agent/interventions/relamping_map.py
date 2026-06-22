"""
Mappa di conversione lampade per l'intervento di relamping.

Contiene:
- Tabella sinonimi tipologia → categoria normalizzata
- Tabelle di conversione P_ante → P_post per categoria (con interpolazione lineare su %)
- Tabella prezzi LED per categoria e P_ante

Fonte: criteri_relamping.md
"""

from __future__ import annotations
import math
from dataclasses import dataclass
from typing import Optional


# ---------------------------------------------------------------------------
# Tabella sinonimi: lista di (categoria, [frammenti testuali])
# Ordine importante: le categorie più specifiche vanno prima di quelle generiche
# ---------------------------------------------------------------------------
_SINONIMI: list[tuple[str, list[str]]] = [
    ("vapori_sodio_bassa_pressione", [
        "vapori sodio bassa pressione",
        "vapore sodio bassa pressione",
    ]),
    ("vapori_mercurio", [
        "vapori di mercurio", "vapore di mercurio", "mercurio",
    ]),
    ("ioduri_metallici_grandi", [
        "ioduri di sodio", "ioduri di sodio",
        "proiettore a sospensione", "proiettore a sospenzione",
        "apparecchi a sospensione",
    ]),
    ("ioduri_metallici_piccoli", [
        "lampada a ioduri metallici", "lampada a ioduri",
        "ioduri metallici", "ioduri",
    ]),
    ("torre_faro", [
        "torre faro", "torri faro",
    ]),
    ("lampione_SAP", [
        "lampione sap", "lampioni sap", "vapori di sodio ap",
        "vapori di sodio alta pressione", "vapore di sodio ap",
        "vapori sodio alta pressione", "lampione", "lampioni",
        "vapori di sodio", "vapore di sodio",
    ]),
    ("faretto_HID", [
        "faretto", "faretti", "faro", "fari",
        "faretti circolari", "faretto circolare", "fari grandi",
        "fari circolari", "faretti agli ioduri metallici",
        "fari ioduri", "faro ioduri", "faretti ioduri", "faro iuduri",
        "fari a ioduri di sodio", "fari disano rodio", "faro gewiss",
        "fari atex", "fari sap", "faro neon", "faro fluorescente",
        "faretto neon", "farretto neon", "proiettore a muro",
    ]),
    ("plafoniera_a_muro", [
        "plafoniera tipo alogena", "applique da parete",
    ]),
    ("plafoniera_neon", [
        "plafoniera neon", "plafoniera circolare", "plafoniera controsoffitto",
        "plafoniera tipo neon", "plafoniera emergenza",
        "lampada singoloneon", "lampada doppioneon",
        "lampione doppioneon", "lampade a fluorescenza",
        "luci a fluorescenza", "fluorescenti compatte",
        "lampada fluorescenza", "tubolare fluorescente",
        "fluorescente tubolare", "fluorescenza", "a fluorescenza",
        "fluorescente", "tubi a fluorescenza", "plafoniera",
    ]),
    ("tubo_neon", [
        "tubo neon", "tubi neon", "bulbo neon", "neon bitubo",
        "paline neon", "palina neon", "neon",
    ]),
    ("alogene", [
        "fari alogeni", "faretti alogeni", "alogena", "alogene",
        "faretti (alogeni?)", "faretto (alogeni?)",
        "vapori di alogenuri", "vapore di alogenuri",
        "piantane", "lanterne", "pannelli",
    ]),
    ("incandescenza", [
        "lampada a incandescenza", "lampade a incandescenza",
        "incandescenza",
    ]),
    ("luci_crepuscolari", [
        "luci crepuscolari", "luci palco", "luce sala attesa", "luci",
    ]),
]


@dataclass
class ConversionPoint:
    p_ante_kw: float
    p_post_kw: float
    riduzione_pct: float
    fonte: str
    is_outlier: bool = False


# Tabelle di conversione per categoria
# Chiave: categoria, Valore: lista di ConversionPoint ordinati per p_ante_kw
_CONVERSION_TABLES: dict[str, list[ConversionPoint]] = {
    "faretto_HID": [
        ConversionPoint(0.020, 0.015, 25.0, "TEAM"),
        ConversionPoint(0.250, 0.150, 40.0, "TEAM"),
        ConversionPoint(0.400, 0.150, 62.0, "RICERCA"),
        ConversionPoint(0.720, 0.310, 57.0, "TEAM"),
        ConversionPoint(1.000, 0.400, 60.0, "RICERCA"),
        ConversionPoint(2.000, 0.500, 75.0, "RICERCA"),
    ],
    "plafoniera_neon": [
        ConversionPoint(0.015, 0.007, 53.0, "TEAM"),
        ConversionPoint(0.018, 0.009, 50.0, "RICERCA"),
        ConversionPoint(0.036, 0.018, 50.0, "RICERCA"),
        ConversionPoint(0.058, 0.028, 52.0, "RICERCA"),
        ConversionPoint(0.072, 0.033, 54.0, "RICERCA"),
        ConversionPoint(0.112, 0.050, 55.0, "RICERCA"),
    ],
    "plafoniera_a_muro": [
        ConversionPoint(0.150, 0.120, 20.0, "TEAM", is_outlier=True),
    ],
    "tubo_neon": [
        ConversionPoint(0.009,  0.0045, 50.0, "TEAM"),
        ConversionPoint(0.018,  0.009,  50.0, "TEAM"),
        ConversionPoint(0.036,  0.018,  50.0, "TEAM"),
        ConversionPoint(0.058,  0.031,  47.0, "TEAM"),
        ConversionPoint(0.096,  0.058,  40.0, "TEAM", is_outlier=True),
    ],
    "torre_faro": [
        ConversionPoint(0.400, 0.220, 45.0, "TEAM"),
        ConversionPoint(1.600, 0.880, 45.0, "TEAM"),
    ],
    "lampione_SAP": [
        ConversionPoint(0.070, 0.040, 43.0, "RICERCA"),
        ConversionPoint(0.100, 0.055, 45.0, "RICERCA"),
        ConversionPoint(0.150, 0.076, 49.0, "TEAM"),
        ConversionPoint(0.250, 0.140, 44.0, "RICERCA"),
        ConversionPoint(0.400, 0.220, 45.0, "RICERCA"),
    ],
    "vapori_sodio_bassa_pressione": [],  # usa default 35%
    "vapori_mercurio": [],               # usa default 55%
    "ioduri_metallici_grandi": [
        ConversionPoint(0.250, 0.100, 60.0, "RICERCA"),
        ConversionPoint(0.400, 0.150, 62.0, "RICERCA"),
        ConversionPoint(1.000, 0.500, 50.0, "RICERCA"),
        ConversionPoint(2.000, 0.600, 70.0, "RICERCA"),
    ],
    "ioduri_metallici_piccoli": [],  # usa default 55%, range 70-150W
    "incandescenza": [
        ConversionPoint(0.040, 0.005, 88.0, "RICERCA"),
        ConversionPoint(0.060, 0.008, 87.0, "RICERCA"),
        ConversionPoint(0.100, 0.015, 85.0, "RICERCA"),
    ],
    "alogene": [
        ConversionPoint(0.050, 0.007, 86.0, "RICERCA"),
        ConversionPoint(0.060, 0.007, 88.0, "RICERCA"),
    ],
    "luci_crepuscolari": [],  # fallback generico
}

# Default % riduzione per categorie senza tabella puntuale
_DEFAULT_RIDUZIONE: dict[str, float] = {
    "vapori_sodio_bassa_pressione": 35.0,
    "vapori_mercurio": 55.0,
    "ioduri_metallici_piccoli": 55.0,
    "incandescenza": 85.0,
    "alogene": 80.0,
    "luci_crepuscolari": 50.0,
}

# Tabella prezzi LED: categoria → [(P_ante_min, P_ante_max, prezzo_euro)]
# Usare il valore massimo della forbice (criterio cautelativo)
_PREZZI: dict[str, list[tuple[float, float, float]]] = {
    "faretto_HID": [
        (0.020, 0.020, 60.0),
        (0.250, 0.250, 160.0),
        (0.400, 0.400, 220.0),
        (0.720, 0.720, 140.0),
        (1.000, 1.000, 350.0),
    ],
    "plafoniera_neon": [
        (0.015, 0.015, 25.0),
        (0.018, 0.058, 25.0),
        (0.072, 0.112, 35.0),
    ],
    "plafoniera_a_muro": [
        (0.150, 0.150, 200.0),
    ],
    "tubo_neon": [
        (0.009, 0.096, 35.0),
    ],
    "torre_faro": [
        (0.400, 0.400, 284.0),
        (1.600, 1.600, 550.0),
    ],
    "lampione_SAP": [
        (0.070, 0.150, 220.0),
        (0.250, 0.400, 350.0),
    ],
    "vapori_mercurio": [
        (0.0, 9999.0, 300.0),
    ],
    "ioduri_metallici_piccoli": [
        (0.070, 0.150, 150.0),
    ],
    "incandescenza": [
        (0.040, 0.100, 30.0),
    ],
    "alogene": [
        (0.050, 2.000, 80.0),
    ],
    "vapori_sodio_bassa_pressione": [
        (0.0, 9999.0, 220.0),  # non in tabella, uso lampione_SAP piccolo come riferimento
    ],
    "ioduri_metallici_grandi": [
        (0.250, 0.400, 220.0),
        (1.000, 2.000, 350.0),
    ],
    "luci_crepuscolari": [
        (0.0, 9999.0, 80.0),
    ],
}

FALLBACK_RIDUZIONE = 50.0
FALLBACK_CATEGORIA = "_fallback"


@dataclass
class ConversionResult:
    categoria: str
    p_post_kw: float
    riduzione_pct: float
    confidenza: str  # "Alta", "Media", "Bassa", "Da verificare"
    costo_lampada_euro: float
    note: list[str]


def normalizza_tipologia(tipologia: str) -> str:
    """
    Mappa una stringa tipologia alla categoria normalizzata.
    I pattern più lunghi hanno priorità su quelli corti (evita che 'ioduri'
    da solo intercetti 'faro ioduri' prima che lo faccia il pattern specifico).
    """
    t = tipologia.lower().strip()
    # Costruisce lista piatta (categoria, sinonimo) ordinata per lunghezza decrescente
    candidati: list[tuple[str, str]] = []
    for categoria, sinonimi in _SINONIMI:
        for s in sinonimi:
            candidati.append((categoria, s))
    candidati.sort(key=lambda x: len(x[1]), reverse=True)
    for categoria, s in candidati:
        if s in t:
            return categoria
    return FALLBACK_CATEGORIA


def _interpola_riduzione(p_kw: float, punti: list[ConversionPoint]) -> tuple[float, str, list[str]]:
    """
    Calcola la % riduzione per p_kw interpolando/estrapolando sulla tabella.
    Ritorna (riduzione_pct, confidenza, note).
    """
    note = []
    if not punti:
        return FALLBACK_RIDUZIONE, "Bassa", ["Tabella vuota, usato fallback 50%"]

    if len(punti) == 1:
        p0 = punti[0]
        if abs(p_kw - p0.p_ante_kw) < 1e-6:
            conf = "Alta" if p0.fonte == "TEAM" else "Media"
        else:
            conf = "Bassa"
            note.append(f"Unico punto in tabella ({p0.p_ante_kw} kW), categoria a basso campione")
        if p0.is_outlier:
            conf = "Da verificare"
            note.append("Valore outlier, verificare in loco")
        return p0.riduzione_pct, conf, note

    # Cerca match esatto
    for pt in punti:
        if abs(p_kw - pt.p_ante_kw) < 1e-6:
            conf = "Alta" if pt.fonte == "TEAM" else "Media"
            if pt.is_outlier:
                conf = "Da verificare"
                note.append("Valore outlier (anomalia rispetto al pattern), verificare in loco")
            return pt.riduzione_pct, conf, note

    # Interpola tra due punti
    sotto = [pt for pt in punti if pt.p_ante_kw < p_kw]
    sopra = [pt for pt in punti if pt.p_ante_kw > p_kw]

    if sotto and sopra:
        p_low = max(sotto, key=lambda x: x.p_ante_kw)
        p_high = min(sopra, key=lambda x: x.p_ante_kw)
        frac = (p_kw - p_low.p_ante_kw) / (p_high.p_ante_kw - p_low.p_ante_kw)
        riduzione = p_low.riduzione_pct + frac * (p_high.riduzione_pct - p_low.riduzione_pct)
        fonte_low = "TEAM" if p_low.fonte == "TEAM" else "RICERCA"
        fonte_high = "TEAM" if p_high.fonte == "TEAM" else "RICERCA"
        conf = "Alta" if fonte_low == "TEAM" and fonte_high == "TEAM" else "Media"
        return riduzione, conf, note

    # Estrapolazione
    if not sotto:
        pt = min(punti, key=lambda x: x.p_ante_kw)
    else:
        pt = max(punti, key=lambda x: x.p_ante_kw)
    note.append(f"Potenza {p_kw} kW fuori dal range tabella, usato limite più vicino ({pt.p_ante_kw} kW) — estrapolato")
    return pt.riduzione_pct, "Bassa", note


def _get_prezzo(categoria: str, p_ante_kw: float) -> float:
    """Restituisce il prezzo LED per la categoria e potenza ante."""
    tabella = _PREZZI.get(categoria, [])
    if not tabella:
        return 80.0  # fallback generico

    # Prima prova match di range
    for p_min, p_max, prezzo in tabella:
        if p_min - 1e-6 <= p_ante_kw <= p_max + 1e-6:
            return prezzo

    # Nearest neighbor per P_ante
    nearest = min(tabella, key=lambda x: abs((x[0] + x[1]) / 2 - p_ante_kw))
    return nearest[2]


def converti_lampada(tipologia: str, p_ante_kw: float) -> ConversionResult:
    """
    Dato il tipo di lampada e la potenza ante, restituisce la potenza LED post
    e il costo della lampada.
    """
    note: list[str] = []

    if not tipologia or not tipologia.strip():
        return ConversionResult(
            categoria=FALLBACK_CATEGORIA,
            p_post_kw=p_ante_kw * (1 - FALLBACK_RIDUZIONE / 100),
            riduzione_pct=FALLBACK_RIDUZIONE,
            confidenza="Bassa",
            costo_lampada_euro=80.0,
            note=["Tipologia vuota, applicato fallback generico 50%"],
        )

    categoria = normalizza_tipologia(tipologia)

    if categoria == FALLBACK_CATEGORIA:
        note.append(f"Tipologia '{tipologia}' non mappata, applicato fallback generico 50%")
        return ConversionResult(
            categoria=FALLBACK_CATEGORIA,
            p_post_kw=p_ante_kw * (1 - FALLBACK_RIDUZIONE / 100),
            riduzione_pct=FALLBACK_RIDUZIONE,
            confidenza="Bassa",
            costo_lampada_euro=80.0,
            note=note,
        )

    # Categorie con solo default (nessuna tabella puntuale)
    if categoria in _DEFAULT_RIDUZIONE and not _CONVERSION_TABLES.get(categoria):
        riduzione = _DEFAULT_RIDUZIONE[categoria]
        note.append(f"Categoria '{categoria}': applicata riduzione default {riduzione}%")
        conf = "Media"
    else:
        punti = _CONVERSION_TABLES.get(categoria, [])
        riduzione, conf, extra_note = _interpola_riduzione(p_ante_kw, punti)
        note.extend(extra_note)

    # Caso speciale torre_faro: % costante 45%, non interpolare verso altre categorie
    if categoria == "torre_faro":
        riduzione = 45.0
        conf = "Alta" if conf != "Bassa" else "Bassa"

    p_post = round(p_ante_kw * (1 - riduzione / 100), 6)
    costo = _get_prezzo(categoria, p_ante_kw)

    return ConversionResult(
        categoria=categoria,
        p_post_kw=p_post,
        riduzione_pct=riduzione,
        confidenza=conf,
        costo_lampada_euro=costo,
        note=note,
    )
