"""
Schema dati canonico per il modello energetico.

Contratto tra il parser Excel e i tool di intervento:
ogni tool riceve un EnergyModel e non sa nulla del formato sorgente.

DA AFFINARE: i campi sono placeholder generici. Completare quando arriva
il file Excel reale con la struttura definitiva delle colonne.
"""

from __future__ import annotations
from typing import Optional, List
from pydantic import BaseModel


class Utenza(BaseModel):
    """
    Singola utenza energetica (edificio, POD, contatore).
    DA AFFINARE: aggiungere/rimuovere campi quando arriva il file reale.
    """
    pod: str
    descrizione: Optional[str] = None
    # DA AFFINARE: unità di misura, fonte dati
    consumo_annuo_kwh: Optional[float] = None
    superficie_mq: Optional[float] = None
    # DA AFFINARE: codifica etichetta energetica (es. "A4", "G")
    etichetta_energetica: Optional[str] = None
    # DA AFFINARE: tipologia edificio (uffici, produzione, magazzino...)
    tipo_edificio: Optional[str] = None
    regione: Optional[str] = None


class EnergyModel(BaseModel):
    """
    Modello energetico completo, prodotto dal parser e consumato dai tool.
    DA AFFINARE: aggiungere metadati di progetto, anno di riferimento, ecc.
    """
    # DA AFFINARE: identificativo progetto/cliente
    nome_progetto: Optional[str] = None
    anno_riferimento: Optional[int] = None
    utenze: List[Utenza] = []
