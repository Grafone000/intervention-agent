"""
Assembla i risultati degli interventi in un documento .docx.

Usa python-docx. Per ora genera un documento minimale (titolo + sezione
per intervento) per verificare che la generazione funzioni nella sandbox.
DA ARRICCHIRE: tabelle riepilogative, grafici inline, frontespizio, ecc.
"""

from pathlib import Path
from typing import List
from docx import Document
from docx.shared import Pt

from interventions.base import InterventionResult


def costruisci_relazione(results: List[InterventionResult], output_path: str) -> str:
    """
    Crea un documento .docx con i risultati degli interventi.

    Args:
        results: lista di InterventionResult da includere nel report
        output_path: percorso di output per il file .docx

    Returns:
        Percorso assoluto del file .docx generato
    """
    doc = Document()

    doc.add_heading("Relazione Benchmark Energetico", level=0)
    doc.add_paragraph(
        "Documento generato automaticamente dall'Energy Benchmark Agent. "
        "Logica di business non ancora implementata — contenuto placeholder."
    )

    for result in results:
        doc.add_heading(result.nome_intervento, level=1)
        p = doc.add_paragraph()
        p.add_run("Risparmio annuo: ").bold = True
        p.add_run(f"{result.risparmio_annuo_kwh:,.0f} kWh/anno\n")
        p.add_run("Costo stimato: ").bold = True
        p.add_run(f"€ {result.costo_stimato_euro:,.0f}\n")
        p.add_run("Payback: ").bold = True
        p.add_run(f"{result.payback_anni:.1f} anni")

        if result.note:
            doc.add_paragraph("Note:")
            for nota in result.note:
                doc.add_paragraph(nota, style="List Bullet")

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return str(out.resolve())
