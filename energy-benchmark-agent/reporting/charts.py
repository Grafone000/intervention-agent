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
    spesa_ante_totale: float,
    spesa_post_totale: float,
) -> io.BytesIO:
    """
    Due barre: spesa totale ante (grigio scuro) e spesa totale post (grigio chiaro).
    """
    fig, ax = plt.subplots(figsize=(5, 5))

    ax.bar([0], [spesa_ante_totale], width=0.4, color="#555555", label="Spesa ante [€/anno]")
    ax.bar([1], [spesa_post_totale], width=0.4, color="#AAAAAA", label="Spesa post [€/anno]")

    ax.set_xticks([0, 1])
    ax.set_xticklabels(["Ante", "Post"], fontsize=10)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"€ {v:,.0f}"))
    ax.set_ylabel("Spesa [€/anno]")
    ax.set_title("Confronto spesa energetica ante/post relamping")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()

    return _buf(fig)


# ─── helpers calendario (anno non bisestile, 8760 ore) ───────────────────────

_GIORNI_MESE = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
_NOMI_MESE = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu",
              "Lug", "Ago", "Set", "Ott", "Nov", "Dic"]


def _ore_inizio_mese() -> List[int]:
    """Ora di inizio (0-based) di ciascun mese in un anno di 8760 ore."""
    inizi = [0]
    for g in _GIORNI_MESE:
        inizi.append(inizi[-1] + g * 24)
    return inizi   # 13 valori: inizi[m]..inizi[m+1] sono le ore del mese m (0-based)


def _aggrega_mensile(serie: List[float]) -> List[float]:
    """Somma i valori orari per ciascuno dei 12 mesi."""
    inizi = _ore_inizio_mese()
    return [sum(serie[inizi[m]:inizi[m + 1]]) for m in range(12)]


def _slice_giorno(serie: List[float], mese: int, giorno: int) -> List[float]:
    """24 valori orari del giorno indicato (mese 1-12, giorno 1-31)."""
    doy = sum(_GIORNI_MESE[:mese - 1]) + (giorno - 1)
    start = doy * 24
    return list(serie[start:start + 24])


def _slice_stagione(serie: List[float], mesi: List[int]) -> List[float]:
    """Concatena le ore dei mesi indicati (mese 1-12)."""
    inizi = _ore_inizio_mese()
    out: List[float] = []
    for m in mesi:
        out.extend(serie[inizi[m - 1]:inizi[m]])
    return out


# ─── Grafici fotovoltaico ────────────────────────────────────────────────────

def grafico_mensile_prod_vs_cons(
    produzione_oraria: List[float],
    consumo_orario: List[float],
) -> io.BytesIO:
    """Barre affiancate per mese: consumi (blu) vs produzione (rosso) [kWh]."""
    prod_m = _aggrega_mensile(produzione_oraria)
    cons_m = _aggrega_mensile(consumo_orario)

    x = list(range(12))
    w = 0.4
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar([i - w / 2 for i in x], cons_m, width=w, color="#1a56db", label="Consumi POD")
    ax.bar([i + w / 2 for i in x], prod_m, width=w, color="#e02424", label="Produzione FV")

    ax.set_xticks(x)
    ax.set_xticklabels(_NOMI_MESE)
    ax.set_ylabel("Energia [kWh]")
    ax.set_title("Confronto mensile produzione FV vs consumi POD")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()
    return _buf(fig)


def grafico_giornata_tipica(
    produzione_oraria: List[float],
    consumo_orario: List[float],
    mese: int,
    giorno: int,
    titolo: str,
) -> io.BytesIO:
    """Due linee (produzione vs consumo) sulle 24 ore di una giornata tipica."""
    prod_d = _slice_giorno(produzione_oraria, mese, giorno)
    cons_d = _slice_giorno(consumo_orario, mese, giorno)
    ore = list(range(24))
    etichette = [f"{h:02d}:10" for h in ore]

    fig, ax = plt.subplots(figsize=(9, 4.5))
    ax.plot(ore, cons_d, color="#1a56db", linewidth=2, marker="o", markersize=3, label="Consumi POD")
    ax.plot(ore, prod_d, color="#e02424", linewidth=2, marker="s", markersize=3, label="Produzione FV")

    ax.set_xticks(ore)
    ax.set_xticklabels(etichette, rotation=90, fontsize=7)
    ax.set_ylabel("Potenza media oraria [kWh]")
    ax.set_title(titolo)
    ax.legend()
    ax.grid(linestyle="--", alpha=0.4)
    fig.tight_layout()
    return _buf(fig)


def grafico_stagionale(
    produzione_oraria: List[float],
    consumo_orario: List[float],
    mesi: List[int],
    titolo: str,
) -> io.BytesIO:
    """Barre sulle ore di una stagione: produzione vs consumo."""
    prod_s = _slice_stagione(produzione_oraria, mesi)
    cons_s = _slice_stagione(consumo_orario, mesi)
    x = list(range(len(prod_s)))

    fig, ax = plt.subplots(figsize=(10, 4.5))
    ax.bar(x, cons_s, width=1.0, color="#1a56db", alpha=0.6, label="Consumi POD")
    ax.bar(x, prod_s, width=1.0, color="#e02424", alpha=0.6, label="Produzione FV")

    ax.set_xlabel("Ore della stagione")
    ax.set_ylabel("Energia oraria [kWh]")
    ax.set_title(titolo)
    ax.set_xlim(0, len(x))
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()
    return _buf(fig)


def grafico_costi_fv(
    costo_ante: float,
    costo_post: float,
    ricavo_immissione: float,
) -> io.BytesIO:
    """Barre: spesa ante, spesa post (residua da rete) e ricavo da immissione."""
    fig, ax = plt.subplots(figsize=(6, 5))
    ax.bar([0], [costo_ante], width=0.5, color="#555555", label="Spesa ante")
    ax.bar([1], [costo_post], width=0.5, color="#AAAAAA", label="Spesa post")
    ax.bar([2], [ricavo_immissione], width=0.5, color="#1f7535", label="Ricavo immissione")

    ax.set_xticks([0, 1, 2])
    ax.set_xticklabels(["Spesa ante", "Spesa post", "Immissione"], fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"€ {v:,.0f}"))
    ax.set_ylabel("€/anno")
    ax.set_title("Spesa energetica ante/post e ricavo da immissione")
    ax.legend()
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    fig.tight_layout()
    return _buf(fig)


def grafico_flussi_cassa_fv(
    fc_attualizzato: List[float],
    fc_semplice: List[float],
    investimento: float = 0.0,
) -> io.BytesIO:
    """Flussi di cassa cumulati FV: attualizzato (blu) e semplice (rosso)."""
    anni = len(fc_attualizzato)
    x_vals = list(range(0, anni + 1))

    def prepend_zero(lst: List[float]) -> List[float]:
        return [-investimento] + list(lst)

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(x_vals, prepend_zero(fc_attualizzato), color="#1a56db", linestyle="-",
            linewidth=2, marker="o", markersize=4, label="FC Attualizzato")
    ax.plot(x_vals, prepend_zero(fc_semplice), color="#e02424", linestyle="-",
            linewidth=2, marker="s", markersize=4, label="FC Semplice")

    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_xlabel("Anno")
    ax.set_ylabel("Flusso di cassa cumulato [€]")
    ax.set_title("Analisi dei flussi di cassa — impianto fotovoltaico")
    ax.set_xticks(x_vals)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"€ {v:,.0f}"))
    ax.legend()
    ax.grid(linestyle="--", alpha=0.4)
    fig.tight_layout()
    return _buf(fig)


def grafico_flussi_cassa(
    fc_att_con: List[float],
    fc_att_senza: List[float],
    fc_sempl_con: List[float],
    fc_sempl_senza: List[float],
    investimento: float = 0.0,
) -> io.BytesIO:
    """
    Grafico lineare dei 4 scenari di flusso di cassa cumulato.
    Blu = attualizzati, Rosso = semplici.
    Solido = con incentivi, Tratteggiato = senza incentivi.
    Le liste in ingresso hanno 8 valori (anni 1-8); viene preposto l'anno 0 = -investimento.
    """
    anni = len(fc_att_con)
    x_vals = list(range(0, anni + 1))

    def prepend_zero(lst: List[float]) -> List[float]:
        return [-investimento] + list(lst)

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
