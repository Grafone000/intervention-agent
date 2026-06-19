"""
Parser Excel -> EnergyModel.

Legge il foglio ModEE e produce un EnergyModel validato.
Gestisce:
  - celle con testo invece di formula (es. h_eq = 'stile ') → ricalcolo automatico
  - celle #VALUE! → ricalcolo automatico dai campi primitivi
  - utenze con Tipologia None → incluse nel modello con note_parser
  - normalizzazione nomi colonna (.strip())

Fattori di conversione (da foglio FATTORI CONVERSIONE):
  tep:  Consumo × 0.000187  (kWh → tep energia elettrica)
  CO2:  Consumo × 294.784 / 1000  (kWh → kg CO2)
  €:    Consumo × 0.21  (€/kWh)
"""

from __future__ import annotations
from pathlib import Path
from typing import Optional
import openpyxl

from .energy_model import EnergyModel, Utenza

# Fattori di conversione da foglio FATTORI CONVERSIONE
_FATTORE_TEP = 0.000187       # C2
_FATTORE_CO2 = 294.784 / 1000 # C15 (kg/MWh → kg/kWh)
_PREZZO_KWH  = 0.21           # C24

# Mappa nomi colonna attesi (dopo .strip()) → indice 0-based
_COL_MAP = {
    "Utenza":                        2,
    "Tipologia":                     3,
    "POD":                           4,
    "Edificio/zona":                 5,
    "Competenza":                   10,
    "Uso energetico":               12,
    "N° di utenze":                 13,
    "Potenza [kW]":                 15,
    "Potenza totale [kW]":          16,
    "Fattore di Utilizzo [%]":      18,
    "Fattore di contemporaneità [%]": 19,
    "h/d":                          20,
    "d/w":                          21,
    "w/y":                          22,
    "h/y":                          23,
    "h equivalenti [h]":            24,
    "Consumo [kWh]":                25,
    "tep":                          27,
    "Costi energetici [€]":         28,
    "CO2 emessa [kg]":              29,
}


def _is_error(val) -> bool:
    """Restituisce True se la cella contiene un errore Excel o testo inatteso."""
    if val is None:
        return False
    if isinstance(val, str):
        return val.strip().startswith("#") or not val.strip().replace(".", "").replace(",", "").lstrip("-").isdigit()
    return False


def _to_float(val, allow_none: bool = False) -> Optional[float]:
    """Converte un valore cella in float; None se errore e allow_none=True."""
    if val is None:
        return None if allow_none else None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        s = val.strip()
        if s.startswith("#") or not s:
            return None
        try:
            return float(s.replace(",", "."))
        except ValueError:
            return None
    return None


def parse_excel(filepath: str | Path, regione: Optional[str] = None) -> EnergyModel:
    """
    Legge il file Excel (foglio ModEE) e restituisce un EnergyModel validato.

    Args:
        filepath: percorso al file .xlsx del modello energetico
        regione: regione del sito (parametro esterno, usato per prezziario FV)

    Returns:
        EnergyModel con tutte le utenze parsate

    Raises:
        FileNotFoundError: se il file non esiste
        ValueError: se il foglio ModEE non è presente o mancano colonne obbligatorie
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"File non trovato: {filepath}")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    if "ModEE" not in wb.sheetnames:
        raise ValueError(f"Foglio 'ModEE' non trovato in {filepath.name}. Fogli disponibili: {wb.sheetnames}")

    ws = wb["ModEE"]

    # Leggi e valida header (riga 1), normalizzando con .strip()
    header_row = [cell.value for cell in ws[1]]
    header_clean = [str(h).strip() if h is not None else None for h in header_row]

    # Verifica colonne obbligatorie
    obbligatorie = ["Utenza", "POD", "Uso energetico", "N° di utenze",
                    "Potenza [kW]", "h/d", "d/w", "w/y",
                    "Fattore di Utilizzo [%]", "Fattore di contemporaneità [%]"]
    for col in obbligatorie:
        idx = _COL_MAP[col]
        found = header_clean[idx] if idx < len(header_clean) else None
        if found != col:
            raise ValueError(
                f"Colonna obbligatoria '{col}' non trovata in posizione attesa (col {idx+1}). "
                f"Trovato: {repr(found)}"
            )

    utenze: list[Utenza] = []
    errori_bloccanti: list[str] = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # Salta righe senza Utenza (separatori/vuote)
        utenza_val = row[_COL_MAP["Utenza"]] if len(row) > _COL_MAP["Utenza"] else None
        if utenza_val is None or str(utenza_val).strip() == "":
            continue

        def get(col_name: str):
            idx = _COL_MAP[col_name]
            return row[idx] if idx < len(row) else None

        note: list[str] = []

        # Campi obbligatori — errore bloccante se mancano
        pod = get("POD")
        if not pod:
            errori_bloccanti.append(f"Riga {row_idx}: POD mancante per utenza '{utenza_val}'")
            continue

        uso = get("Uso energetico")
        if not uso:
            errori_bloccanti.append(f"Riga {row_idx}: Uso energetico mancante per utenza '{utenza_val}'")
            continue

        # Campi numerici primitivi (base per ricalcoli)
        n_utenze_raw   = _to_float(get("N° di utenze"))
        pot_unit_raw   = _to_float(get("Potenza [kW]"))
        pot_tot_raw    = _to_float(get("Potenza totale [kW]"))
        f_util_raw     = _to_float(get("Fattore di Utilizzo [%]"))
        f_cont_raw     = _to_float(get("Fattore di contemporaneità [%]"))
        hd_raw         = _to_float(get("h/d"))
        dw_raw         = _to_float(get("d/w"))
        wy_raw         = _to_float(get("w/y"))

        # Verifica campi primitivi obbligatori
        missing = [n for n, v in [
            ("N° di utenze", n_utenze_raw), ("Potenza [kW]", pot_unit_raw),
            ("Fattore di Utilizzo [%]", f_util_raw), ("Fattore di contemporaneità [%]", f_cont_raw),
            ("h/d", hd_raw), ("d/w", dw_raw), ("w/y", wy_raw)
        ] if v is None]
        if missing:
            errori_bloccanti.append(
                f"Riga {row_idx}: campi obbligatori mancanti per '{utenza_val}': {missing}"
            )
            continue

        # Potenza totale: usa il valore dal file se disponibile, altrimenti ricalcola
        # Formula: Pot_tot = P * N / COP (COP = col O; per illuminazione COP=1)
        cop = _to_float(get("COP/EER") if "COP/EER" in _COL_MAP else None) or 1.0
        if pot_tot_raw is None:
            pot_tot = pot_unit_raw * n_utenze_raw / cop
            note.append(f"Potenza totale ricalcolata ({pot_unit_raw}×{n_utenze_raw}/{cop})")
        else:
            pot_tot = pot_tot_raw

        # h/y — sempre ricalcolato dai primitivi (più affidabile della cella)
        ore_anno = hd_raw * dw_raw * wy_raw

        # h_eq — se la cella è errata (testo/None) ricalcola; altrimenti usa il valore
        h_eq_raw = get("h equivalenti [h]")
        h_eq_val = _to_float(h_eq_raw)
        if h_eq_val is None:
            h_eq = ore_anno * f_util_raw
            note.append(
                f"h_eq ricalcolato (cella Y{row_idx} conteneva {repr(h_eq_raw)}): "
                f"{ore_anno} × {f_util_raw} = {h_eq:.2f}"
            )
        else:
            h_eq = h_eq_val

        # Consumo — ricalcola se #VALUE! o None
        consumo_raw = get("Consumo [kWh]")
        consumo_val = _to_float(consumo_raw)
        if consumo_val is None:
            consumo = pot_tot * h_eq * f_cont_raw
            note.append(f"Consumo ricalcolato (cella Z{row_idx} era errore): {consumo:.2f} kWh")
        else:
            consumo = consumo_val

        # tep, Costi, CO2 — ricalcola sempre dai fattori di conversione
        # (più robusto che leggere dal file, che dipende dal foglio FATTORI CONVERSIONE)
        tep_raw   = get("tep")
        costi_raw = get("Costi energetici [€]")
        co2_raw   = get("CO2 emessa [kg]")

        tep_val   = _to_float(tep_raw)
        costi_val = _to_float(costi_raw)
        co2_val   = _to_float(co2_raw)

        if tep_val is None:
            tep_val = consumo * _FATTORE_TEP
            note.append(f"tep ricalcolato da Consumo (cella AB{row_idx} era errore)")
        if costi_val is None:
            costi_val = consumo * _PREZZO_KWH
            note.append(f"Costi ricalcolati da Consumo (cella AC{row_idx} era errore)")
        if co2_val is None:
            co2_val = consumo * _FATTORE_CO2
            note.append(f"CO2 ricalcolata da Consumo (cella AD{row_idx} era errore)")

        # Tipologia: warning se mancante (non bloccante)
        tipologia = get("Tipologia")
        if tipologia is not None:
            tipologia = str(tipologia).strip() or None
        if tipologia is None and str(uso).strip() == "Illuminazione":
            note.append(f"Tipologia mancante (cella D{row_idx}): utenza esclusa dal relamping")

        utenze.append(Utenza(
            utenza=str(utenza_val).strip(),
            tipologia=tipologia,
            pod=str(pod).strip(),
            edificio_zona=str(get("Edificio/zona") or "").strip(),
            competenza=str(get("Competenza") or "").strip(),
            uso_energetico=str(uso).strip(),
            n_utenze=int(n_utenze_raw),
            potenza_unitaria_kw=pot_unit_raw,
            potenza_totale_kw=pot_tot,
            fatt_utilizzo=f_util_raw,
            fatt_contemporaneita=f_cont_raw,
            ore_giorno=hd_raw,
            giorni_settimana=dw_raw,
            settimane_anno=wy_raw,
            ore_anno=ore_anno,
            h_equivalenti=h_eq,
            consumo_kwh=consumo,
            tep=tep_val,
            costi_euro=costi_val,
            co2_kg=co2_val,
            note_parser=note,
        ))

    if errori_bloccanti:
        raise ValueError(
            f"Parsing fallito: {len(errori_bloccanti)} errori bloccanti:\n" +
            "\n".join(f"  - {e}" for e in errori_bloccanti)
        )

    return EnergyModel(
        nome_progetto=filepath.stem,
        regione=regione,
        utenze=utenze,
    )
