"""
Lookup prezzi unitari per intervento energetico.

I prezzi sono letti da CSV in pricing/prezziari/<regione>.csv.
Struttura attesa del CSV: colonne [intervento, anno, prezzo_euro_unita]
DA AFFINARE quando arrivano i prezziari reali.
"""

from pathlib import Path
from typing import Optional
import pandas as pd

_PREZZIARI_DIR = Path(__file__).parent / "prezziari"


def get_prezzo(intervento: str, regione: str, anno: Optional[int] = None) -> float:
    """
    Restituisce il prezzo unitario per un intervento in una regione.

    Args:
        intervento: nome intervento (deve corrispondere a NOME_INTERVENTO)
        regione: regione italiana (es. "lombardia")
        anno: anno di riferimento; se None usa l'anno più recente disponibile

    Returns:
        Prezzo in euro/unità (unità dipende dall'intervento)

    Raises:
        FileNotFoundError: se non esiste un CSV per la regione indicata
        KeyError: se l'intervento non è presente nel CSV
        ValueError: se l'anno non è disponibile nel CSV
    """
    csv_path = _PREZZIARI_DIR / f"{regione.lower()}.csv"
    if not csv_path.exists():
        raise FileNotFoundError(
            f"Nessun prezziario trovato per la regione '{regione}'. "
            f"File atteso: {csv_path}"
        )

    df = pd.read_csv(csv_path)

    if "intervento" not in df.columns or "prezzo_euro_unita" not in df.columns:
        raise ValueError(
            f"Il CSV {csv_path} non ha le colonne attese "
            "['intervento', 'anno', 'prezzo_euro_unita']."
        )

    df_int = df[df["intervento"] == intervento]
    if df_int.empty:
        raise KeyError(
            f"Intervento '{intervento}' non trovato nel prezziario di '{regione}'."
        )

    if anno is not None and "anno" in df.columns:
        df_int = df_int[df_int["anno"] == anno]
        if df_int.empty:
            raise ValueError(
                f"Anno {anno} non disponibile per '{intervento}' in '{regione}'."
            )
    elif "anno" in df.columns:
        # Usa anno più recente disponibile
        df_int = df_int[df_int["anno"] == df_int["anno"].max()]

    return float(df_int["prezzo_euro_unita"].iloc[0])
