"""
Parser file consumi energetici → array orario (8760 valori) [kWh].

Formati supportati:

1. **Formato standard quart'orario** (CSV o Excel):
   Colonne: timestamp | consumo_kwh (un POD, 35.040 righe/anno)
   → somma 4 quarter per ora → 8760 valori

2. **Formato multi-POD orario** (Excel — formato interno):
   Foglio: "Input POD orario ATTIVA"
   Struttura: righe = ore (8760), colonne POD dalla colonna DZ (130) in poi
   → somma tutti i POD per ogni ora → 8760 valori
   Usato da parse_consumi_excel_multi_pod()
"""

from __future__ import annotations
from pathlib import Path
from typing import List, Optional

import csv


def _norm_pod(s: object) -> str:
    """Normalizza un codice POD per il confronto (strip + upper)."""
    return str(s).strip().upper()


def lista_pod_disponibili(
    filepath: str | Path,
    nome_foglio: str = "Input POD orario ATTIVA",
    col_prima_pod: int = 130,
) -> List[str]:
    """
    Restituisce l'elenco dei codici POD presenti nel foglio multi-POD.

    I codici sono letti dalla riga di intestazione (prima riga) a partire dalla
    colonna `col_prima_pod` (1-based, default 130 = colonna DZ).

    Returns:
        Lista di codici POD nell'ordine in cui compaiono nel foglio.
    """
    import openpyxl
    fp = Path(filepath)
    if not fp.exists():
        raise FileNotFoundError(f"File consumi non trovato: {fp}")

    wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
    if nome_foglio not in wb.sheetnames:
        raise ValueError(
            f"Foglio '{nome_foglio}' non trovato. Fogli disponibili: {wb.sheetnames}"
        )
    ws = wb[nome_foglio]

    header_row = next(ws.iter_rows(min_col=col_prima_pod, values_only=True), None)
    if header_row is None:
        return []

    pod: List[str] = []
    for v in header_row:
        if v is None:
            continue
        code = str(v).strip()
        if code:
            pod.append(code)
    return pod


def _leggi_colonna_consumi_csv(filepath: Path) -> List[float]:
    """Legge un CSV e restituisce la colonna dei consumi come lista di float."""
    valori: List[float] = []
    with open(filepath, newline="", encoding="utf-8-sig") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        reader = csv.reader(f, dialect)

        header = next(reader, None)
        if header is None:
            raise ValueError(f"File CSV vuoto: {filepath}")

        # Trova indice colonna consumi: cerca 'consumo' nel header, fallback colonna 1
        idx_consumo = 1
        if header:
            for i, h in enumerate(header):
                if "consumo" in h.lower() or "kwh" in h.lower() or "energia" in h.lower():
                    idx_consumo = i
                    break

        for row in reader:
            if not row or len(row) <= idx_consumo:
                continue
            raw = row[idx_consumo].strip().replace(",", ".")
            if not raw:
                continue
            try:
                valori.append(float(raw))
            except ValueError:
                continue

    return valori


def _leggi_colonna_consumi_excel(filepath: Path) -> List[float]:
    """Legge un Excel e restituisce la colonna dei consumi come lista di float."""
    try:
        import openpyxl
    except ImportError:
        raise ImportError("openpyxl richiesto per leggere file Excel")

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"File Excel vuoto: {filepath}")

    header = [str(c).strip().lower() if c else "" for c in rows[0]]
    idx_consumo = 1
    for i, h in enumerate(header):
        if "consumo" in h or "kwh" in h or "energia" in h:
            idx_consumo = i
            break

    valori: List[float] = []
    for row in rows[1:]:
        if not row or len(row) <= idx_consumo:
            continue
        val = row[idx_consumo]
        if val is None:
            continue
        try:
            valori.append(float(val))
        except (TypeError, ValueError):
            continue

    return valori


def parse_consumi_excel_multi_pod(
    filepath: str | Path,
    nome_foglio: str = "Input POD orario ATTIVA",
    col_prima_pod: int = 130,   # colonna DZ = indice 0-based 129, ma iter_rows è 1-based → 130
    pod_selezionati: Optional[List[str]] = None,
) -> List[float]:
    """
    Legge il file consumi nel formato multi-POD orario interno.

    Foglio "Input POD orario ATTIVA":
    - Righe = ore dell'anno (8760 righe dati, precedute da header)
    - Colonne POD dalla colonna DZ (130) in poi; la riga di header contiene i codici POD
    - Somma per ogni ora i POD richiesti → array di 8760 valori [kWh]

    Args:
        filepath: percorso al file Excel
        nome_foglio: nome del foglio (default "Input POD orario ATTIVA")
        col_prima_pod: indice 1-based della prima colonna POD (default 130 = colonna DZ)
        pod_selezionati: elenco di codici POD da sommare. Se None o vuoto, somma TUTTI i POD.

    Returns:
        Lista di 8760 valori [kWh/h] — somma dei soli POD richiesti

    Raises:
        ValueError: se nessuno dei POD richiesti è presente nel foglio
    """
    import openpyxl
    fp = Path(filepath)
    if not fp.exists():
        raise FileNotFoundError(f"File consumi non trovato: {fp}")

    wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
    if nome_foglio not in wb.sheetnames:
        raise ValueError(
            f"Foglio '{nome_foglio}' non trovato. Fogli disponibili: {wb.sheetnames}"
        )
    ws = wb[nome_foglio]

    # Insieme dei POD richiesti (normalizzati); None → tutti
    richiesti = {_norm_pod(p) for p in pod_selezionati} if pod_selezionati else None

    orari: List[float] = []
    header_skipped = False
    col_mask: Optional[List[bool]] = None   # quali colonne sommare (allineate a `row`)

    for row in ws.iter_rows(min_col=col_prima_pod, values_only=True):
        # La prima riga è l'header con i codici POD
        if not header_skipped:
            header_skipped = True
            if richiesti is not None:
                col_mask = [
                    (v is not None and _norm_pod(v) in richiesti) for v in row
                ]
                trovati = {
                    _norm_pod(v) for v, keep in zip(row, col_mask) if keep
                }
                mancanti = richiesti - trovati
                if not trovati:
                    disponibili = [str(v).strip() for v in row if v not in (None, "")]
                    raise ValueError(
                        f"Nessuno dei POD richiesti {sorted(richiesti)} è presente nel "
                        f"foglio '{nome_foglio}'. POD disponibili: {disponibili}"
                    )
                if mancanti:
                    # Avviso non bloccante: si procede con i POD trovati
                    import warnings
                    warnings.warn(
                        f"POD non trovati e ignorati: {sorted(mancanti)}", stacklevel=2
                    )
            continue

        # Somma i valori numerici delle colonne selezionate (o tutte)
        totale_ora = 0.0
        has_data = False
        for j, v in enumerate(row):
            if v is None:
                continue
            if col_mask is not None and not (j < len(col_mask) and col_mask[j]):
                continue
            try:
                totale_ora += float(v)
                has_data = True
            except (TypeError, ValueError):
                continue

        if has_data:
            orari.append(totale_ora)

        if len(orari) >= 8760:
            break

    if not orari:
        raise ValueError(
            f"Nessun dato trovato nel foglio '{nome_foglio}' "
            f"dalla colonna {col_prima_pod} in poi."
        )

    # Porta a esattamente 8760 ore
    if len(orari) < 8760:
        orari.extend([0.0] * (8760 - len(orari)))

    return orari[:8760]


def parse_consumi_quart_orari(filepath: str | Path) -> List[float]:
    """
    Legge un file di consumi quart'orari e restituisce array orario [kWh].

    Args:
        filepath: percorso al file CSV o Excel con consumi quart'orari

    Returns:
        Lista di 8760 valori [kWh/h] (consumi orari aggregati)
        Se il file ha meno di 8760 righe i valori mancanti sono 0.

    Raises:
        FileNotFoundError: se il file non esiste
        ValueError: se il file è vuoto o il formato non è riconoscibile
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"File consumi non trovato: {filepath}")

    ext = filepath.suffix.lower()
    if ext in (".xlsx", ".xls"):
        quart_orari = _leggi_colonna_consumi_excel(filepath)
    else:
        quart_orari = _leggi_colonna_consumi_csv(filepath)

    if not quart_orari:
        raise ValueError(f"Nessun dato di consumo trovato in {filepath}")

    # Aggrega 4 quarter → 1 ora
    n_ore = len(quart_orari) // 4
    orari: List[float] = []
    for h in range(n_ore):
        orari.append(sum(quart_orari[h * 4: h * 4 + 4]))

    # Porta a esattamente 8760 ore
    if len(orari) < 8760:
        orari.extend([0.0] * (8760 - len(orari)))
    else:
        orari = orari[:8760]

    return orari
