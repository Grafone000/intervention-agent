"""
Interfaccia base e schema risultato per tutti i tool di intervento.

Ogni intervento deve:
1. Esporre NOME_INTERVENTO: str
2. Esporre una funzione calcola(model, parametri) -> InterventionResult
3. Rispettare il protocollo BaseIntervention (verificato a runtime dal registry)
"""

from __future__ import annotations
from typing import Any, Dict, List, Protocol, runtime_checkable
from pydantic import BaseModel

from core.energy_model import EnergyModel


class InterventionResult(BaseModel):
    """Risultato standardizzato di un tool di intervento energetico."""
    nome_intervento: str
    risparmio_annuo_kwh: float
    costo_stimato_euro: float
    payback_anni: float
    # Struttura libera: i dati da passare alla funzione genera_grafico()
    dati_grafico: Dict[str, Any] = {}
    # Warning su dati mancanti, assunzioni usate, ecc.
    note: List[str] = []


@runtime_checkable
class BaseIntervention(Protocol):
    """
    Protocollo che ogni modulo intervento deve rispettare.
    Non è necessario ereditare da questa classe — basta implementare calcola().
    """
    NOME_INTERVENTO: str

    def calcola(self, model: EnergyModel, parametri: Dict[str, Any]) -> InterventionResult:
        ...
