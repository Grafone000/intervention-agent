"""
Parser Excel -> EnergyModel.

Responsabilità: leggere il file Excel del modello energetico e produrre
un oggetto EnergyModel validato da Pydantic.

DA IMPLEMENTARE nel prossimo step, quando sarà disponibile il file Excel reale.
Struttura attesa dell'Excel: DA DEFINIRE (colonne, fogli, convenzioni di naming).
"""

from pathlib import Path
from .energy_model import EnergyModel


def parse_excel(filepath: str | Path) -> EnergyModel:
    """
    Legge il file Excel e restituisce un EnergyModel.

    Args:
        filepath: percorso al file .xlsx del modello energetico

    Returns:
        EnergyModel validato

    Raises:
        NotImplementedError: finché non viene implementato
        FileNotFoundError: se il file non esiste
        ValueError: se la struttura Excel non è quella attesa
    """
    raise NotImplementedError(
        "parse_excel() non ancora implementato. "
        "Condividi il file Excel reale per procedere con l'implementazione."
    )
