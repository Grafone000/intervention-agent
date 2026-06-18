"""
Generazione grafici da InterventionResult.

Usa matplotlib per compatibilità con python-docx (immagini PNG).
"""

from pathlib import Path
import matplotlib.pyplot as plt
from interventions.base import InterventionResult


def genera_grafico(result: InterventionResult, output_path: str) -> str:
    """
    Genera un grafico a barre per un InterventionResult e lo salva come PNG.

    Per ora crea un grafico placeholder con dati fittizi per verificare
    che matplotlib funzioni correttamente nella sandbox.
    DA AFFINARE: usare result.dati_grafico per dati reali.

    Args:
        result: risultato dell'intervento
        output_path: percorso di output per il file PNG

    Returns:
        Percorso assoluto del file PNG generato
    """
    fig, ax = plt.subplots(figsize=(8, 4))

    # Placeholder: grafico con i tre KPI principali
    labels = ["Risparmio\n(MWh/anno)", "Costo\n(k€)", "Payback\n(anni)"]
    values = [
        result.risparmio_annuo_kwh / 1000,
        result.costo_stimato_euro / 1000,
        result.payback_anni,
    ]
    bars = ax.bar(labels, values, color=["#2ecc71", "#e74c3c", "#3498db"])
    ax.set_title(f"KPI intervento: {result.nome_intervento}", fontsize=13)
    ax.bar_label(bars, fmt="%.1f", padding=3)
    ax.set_ylabel("Valore")
    fig.tight_layout()

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, dpi=150)
    plt.close(fig)
    return str(out.resolve())
