"""
Generazione grafici per la relazione energetica.
Restituisce oggetti BytesIO pronti per l'embedding in python-docx.
"""

from __future__ import annotations
import io
from typing import List

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker


def _buf(fig: plt.Figure) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def grafico_spesa_ante_post(
    labels: List[str],
    spese_ante: List[float],
    spese_post: List[float],
) -> io.BytesIO:
    """
    Barre verticali grigie: Spesa ante (grigio scuro) vs Spesa post (grigio chiaro)
    raggruppate per fabbricato/tipologia.
    """
    x = list(range(len(labels)))
    width = 0.35

    fig, ax = plt.subplots(figsize=(max(8, len(labels) * 0.9), 5))

    ax.bar([i - width / 2 for i in x], spese_ante, width,
           color="#555555", label="Spesa ante [€/anno]")
    ax.bar([i + width / 2 for i in x], spese_post, width,
           color="#AAAAAA", label="Spesa post [€/anno]")

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=30, ha="right", fontsize=8)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"€ {v:,.0f}"))
    ax.set_ylabel("Spesa [€/anno]")
    ax.set_title("Confronto spesa energetica ante/post relamping")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()

    return _buf(fig)


def grafico_flussi_cassa(
    fc_att_con: List[float],
    fc_att_senza: List[float],
    fc_sempl_con: List[float],
    fc_sempl_senza: List[float],
) -> io.BytesIO:
    """
    Grafico lineare dei 4 scenari di flusso di cassa cumulato.
    Blu = attualizzati, Rosso = semplici.
    Solido = con incentivi, Tratteggiato = senza incentivi.
    Le liste in ingresso hanno 8 valori (anni 1-8); viene preposto l'anno 0 = 0.
    """
    anni = len(fc_att_con)
    x_vals = list(range(0, anni + 1))

    def prepend_zero(lst: List[float]) -> List[float]:
        return [0.0] + list(lst)

    fig, ax = plt.subplots(figsize=(9, 5))

    ax.plot(x_vals, prepend_zero(fc_att_con),   color="#1a56db", linestyle="-",  linewidth=2,
            marker="o", markersize=4, label="FC Attualizzato con incentivi")
    ax.plot(x_vals, prepend_zero(fc_att_senza), color="#1a56db", linestyle="--", linewidth=2,
            marker="o", markersize=4, label="FC Attualizzato senza incentivi")
    ax.plot(x_vals, prepend_zero(fc_sempl_con), color="#e02424", linestyle="-",  linewidth=2,
            marker="s", markersize=4, label="FC Semplice con incentivi")
    ax.plot(x_vals, prepend_zero(fc_sempl_senza), color="#e02424", linestyle="--", linewidth=2,
            marker="s", markersize=4, label="FC Semplice senza incentivi")

    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_xlabel("Anno")
    ax.set_ylabel("Flusso di cassa cumulato [€]")
    ax.set_title("Analisi dei flussi di cassa — 4 scenari")
    ax.set_xticks(x_vals)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"€ {v:,.0f}"))
    ax.legend(fontsize=8)
    ax.grid(linestyle="--", alpha=0.4)
    fig.tight_layout()

    return _buf(fig)
