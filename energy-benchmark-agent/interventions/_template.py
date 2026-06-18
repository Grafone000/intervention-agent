"""
TEMPLATE per un nuovo intervento energetico.
==============================================
Per aggiungere un nuovo intervento:

1. Copia questo file e rinominalo (es. fotovoltaico.py, relamping.py)
2. Imposta NOME_INTERVENTO con il nome univoco dell'intervento
3. Implementa la funzione calcola() con la logica reale
4. Il registry.py la troverà automaticamente — nessun altro file va modificato.

Campi di EnergyModel tipicamente necessari:
- model.utenze[i].consumo_annuo_kwh
- model.utenze[i].superficie_mq
- model.utenze[i].regione  (per il lookup prezzi)
- ... (da specificare per ogni intervento)
"""

from typing import Any, Dict
from core.energy_model import EnergyModel
from interventions.base import InterventionResult

# Identificativo univoco dell'intervento — usato come chiave nel registry
NOME_INTERVENTO: str = "nome_intervento_qui"  # DA MODIFICARE


def calcola(model: EnergyModel, parametri: Dict[str, Any]) -> InterventionResult:
    """
    Calcola il benchmark per questo intervento.

    Args:
        model: modello energetico completo (da core/parser.py)
        parametri: parametri specifici dell'intervento (es. potenza_kwp per FV)

    Returns:
        InterventionResult con risparmio, costo, payback e dati per grafico

    Steps da implementare:
        1. Estrarre i campi rilevanti da model.utenze
        2. Chiamare pricing.lookup.get_prezzo() per il costo unitario
        3. Calcolare risparmio_annuo_kwh con la formula specifica
        4. Calcolare costo_stimato_euro
        5. Calcolare payback_anni = costo / (risparmio * prezzo_energia)
        6. Popolare dati_grafico per reporting/charts.py
        7. Aggiungere note[] per assunzioni o dati mancanti
    """
    raise NotImplementedError(
        f"calcola() non ancora implementato per '{NOME_INTERVENTO}'. "
        "Vedi i commenti in interventions/_template.py per le istruzioni."
    )
