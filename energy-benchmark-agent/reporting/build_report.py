"""
Assembla i risultati degli interventi in un documento .docx.

Struttura per il relamping:
  1. Paragrafo introduttivo
  2. Tabella 1 – Valutazione relamping (per fabbricato/tipologia, 2 righe: kW + kWh)
  3. Tabella 2 – Valutazione energetica e economica
  4. Grafico a barre spesa ante/post
  5. Tabella investimento
  6. Testo Certificati Bianchi
  7. Tabelle benchmark (dati ingresso + risultati con/senza incentivi)
  8. Grafico flussi di cassa
"""

from __future__ import annotations
from collections import defaultdict
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT

from interventions.base import InterventionResult
from interventions.relamping import RelampingResult
from interventions.fotovoltaico import FVResult
from reporting.charts import (
    grafico_spesa_ante_post, grafico_flussi_cassa,
    grafico_mensile_prod_vs_cons, grafico_giornata_tipica, grafico_stagionale,
    grafico_costi_fv, grafico_flussi_cassa_fv,
)


# ─── colori ──────────────────────────────────────────────────────────────────

_VERDE_SCURO  = RGBColor(0x1F, 0x75, 0x35)   # intestazione tabella
_VERDE_CHIARO = RGBColor(0xE2, 0xEF, 0xDA)   # righe pari
_BIANCO       = RGBColor(0xFF, 0xFF, 0xFF)
_GRIGIO_INT   = RGBColor(0xBF, 0xBF, 0xBF)   # intestazioni tabelle secondarie


# ─── helpers XML ─────────────────────────────────────────────────────────────

def _set_cell_bg(cell, rgb: RGBColor) -> None:
    """Imposta il colore di sfondo di una cella."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    hex_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_border(cell, **edges) -> None:
    """
    Imposta i bordi di una cella.
    edges: top, bottom, left, right → dict con 'sz', 'color', 'val'
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side, attrs in edges.items():
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), attrs.get("val", "single"))
        border.set(qn("w:sz"), str(attrs.get("sz", 4)))
        border.set(qn("w:color"), attrs.get("color", "auto"))
        tcBorders.append(border)
    tcPr.append(tcBorders)


def _merge_v(cell_top, cell_bottom) -> None:
    """Unione verticale (vMerge) di due celle nella stessa colonna."""
    tc_top = cell_top._tc
    tcPr_top = tc_top.get_or_add_tcPr()
    vMerge_top = OxmlElement("w:vMerge")
    vMerge_top.set(qn("w:val"), "restart")
    tcPr_top.append(vMerge_top)

    tc_bot = cell_bottom._tc
    tcPr_bot = tc_bot.get_or_add_tcPr()
    vMerge_bot = OxmlElement("w:vMerge")
    tcPr_bot.append(vMerge_bot)


def _bold(run) -> None:
    run.bold = True


def _cell_text(cell, text: str, bold: bool = False, size: int = 9,
               color: Optional[RGBColor] = None, center: bool = False) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color


def _fmt(value: float, decimali: int = 0, prefix: str = "") -> str:
    fmt = f"{value:,.{decimali}f}"
    # swap thousand/decimal separators (IT style)
    fmt = fmt.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{prefix}{fmt}"


# ─── intestazione tabella (riga verde scuro, testo bianco) ───────────────────

def _header_row(table, labels: List[str], font_size: int = 9) -> None:
    row = table.rows[0]
    for i, label in enumerate(labels):
        cell = row.cells[i]
        _set_cell_bg(cell, _VERDE_SCURO)
        _cell_text(cell, label, bold=True, size=font_size,
                   color=_BIANCO, center=True)


# ─── Tabella 1 — Valutazione relamping ───────────────────────────────────────

def _tabella_valutazione_relamping(doc: Document, result: RelampingResult) -> None:
    doc.add_heading("Valutazione relamping", level=2)

    # Aggrega tutte le righe per tipologia (ignora l'edificio)
    per_tipo: dict[str, list] = defaultdict(list)
    for r in result.righe:
        per_tipo[r.tipologia_originale].append(r)

    # Colonne: Tipo di lampada | Caratterizzazione | Ex-ante | Ex-post | N. lampade
    cols = ["Tipo di lampada", "Caratterizzazione", "Ex-ante", "Ex-post", "N. lampade"]

    tipo_list = sorted(per_tipo.keys())
    n_data_rows = len(tipo_list) * 2  # 2 righe per tipologia: kW + kWh
    table = doc.add_table(rows=1 + n_data_rows, cols=len(cols))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    _header_row(table, cols)

    row_idx = 1
    for ti, tipo in enumerate(tipo_list):
        righe_tipo = per_tipo[tipo]
        bg = _VERDE_CHIARO if ti % 2 == 0 else _BIANCO

        n_tot      = sum(x.n_utenze for x in righe_tipo)
        p_ante_tot = sum(x.p_tot_ante_kw for x in righe_tipo)
        p_post_tot = sum(x.p_tot_post_kw for x in righe_tipo)
        c_ante_tot = sum(x.consumo_ante_kwh for x in righe_tipo)
        c_post_tot = sum(x.consumo_post_kwh for x in righe_tipo)

        # riga kW
        row_kw = table.rows[row_idx]
        for c in row_kw.cells:
            _set_cell_bg(c, bg)
        _cell_text(row_kw.cells[0], tipo, bold=False, size=8)
        _cell_text(row_kw.cells[1], "Potenza [kW]", bold=False, size=8, center=True)
        _cell_text(row_kw.cells[2], _fmt(p_ante_tot, 3), size=8, center=True)
        _cell_text(row_kw.cells[3], _fmt(p_post_tot, 3), size=8, center=True)
        _cell_text(row_kw.cells[4], str(n_tot), size=8, center=True)

        # riga kWh
        row_kwh = table.rows[row_idx + 1]
        for c in row_kwh.cells:
            _set_cell_bg(c, bg)
        _cell_text(row_kwh.cells[0], tipo, bold=False, size=8)
        _cell_text(row_kwh.cells[1], "Consumi [kWh]", bold=False, size=8, center=True)
        _cell_text(row_kwh.cells[2], _fmt(c_ante_tot, 0), size=8, center=True)
        _cell_text(row_kwh.cells[3], _fmt(c_post_tot, 0), size=8, center=True)
        _cell_text(row_kwh.cells[4], str(n_tot), size=8, center=True)

        # Unione verticale col 0 (tipo) e col 4 (n_lampade) sulle 2 righe
        _merge_v(row_kw.cells[0], row_kwh.cells[0])
        _merge_v(row_kw.cells[4], row_kwh.cells[4])

        row_idx += 2


# ─── Tabella 2 — Valutazione energetica e economica ──────────────────────────

def _tabella_economica(doc: Document, result: RelampingResult) -> None:
    doc.add_heading("Valutazione energetica e economica", level=2)

    cols = [
        "Fabbricato", "Tipo di lampada",
        "Spesa ante\n[€/anno]", "Spesa post\n[€/anno]",
        "Risparmio EE\n[kWh/anno]", "Risparmio\n[€/anno]",
        "Risparmio CO₂\n[kg/anno]",
    ]

    gruppi: dict[str, dict[str, list]] = defaultdict(lambda: defaultdict(list))
    for r in result.righe:
        gruppi[r.zona][r.tipologia_originale].append(r)

    zona_list = sorted(gruppi.keys())
    n_rows = sum(len(t) for t in gruppi.values())

    table = doc.add_table(rows=1 + n_rows + 1, cols=len(cols))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    _header_row(table, cols)

    row_idx = 1
    totale_spesa_ante = 0.0
    totale_spesa_post = 0.0
    totale_risp_ee = 0.0
    totale_risp_euro = 0.0
    totale_risp_co2 = 0.0

    for zi, zona in enumerate(zona_list):
        tipologie_dict = gruppi[zona]
        tipo_list = sorted(tipologie_dict.keys())
        zona_bg = _VERDE_CHIARO if zi % 2 == 0 else _BIANCO
        fabbricato_start = row_idx

        for tipo in tipo_list:
            righe_tipo = tipologie_dict[tipo]
            sp_ante = sum(x.spesa_ante_euro for x in righe_tipo)
            sp_post = sum(x.spesa_post_euro for x in righe_tipo)
            rs_ee   = sum(x.risparmio_ee_kwh for x in righe_tipo)
            rs_eur  = sum(x.risparmio_euro for x in righe_tipo)
            rs_co2  = sum(x.risparmio_co2_kg for x in righe_tipo)

            totale_spesa_ante += sp_ante
            totale_spesa_post += sp_post
            totale_risp_ee   += rs_ee
            totale_risp_euro += rs_eur
            totale_risp_co2  += rs_co2

            row = table.rows[row_idx]
            for c in row.cells:
                _set_cell_bg(c, zona_bg)

            _cell_text(row.cells[0], zona, bold=True, size=8, center=True)
            _cell_text(row.cells[1], tipo, size=8)
            _cell_text(row.cells[2], _fmt(sp_ante, 2, "€ "), size=8, center=True)
            _cell_text(row.cells[3], _fmt(sp_post, 2, "€ "), size=8, center=True)
            _cell_text(row.cells[4], _fmt(rs_ee, 0), size=8, center=True)
            _cell_text(row.cells[5], _fmt(rs_eur, 2, "€ "), size=8, center=True)
            _cell_text(row.cells[6], _fmt(rs_co2, 1), size=8, center=True)
            row_idx += 1

        # merge fabbricato column
        tc0 = table.rows[fabbricato_start].cells[0]._tc
        tcPr0 = tc0.get_or_add_tcPr()
        vMerge_start = OxmlElement("w:vMerge")
        vMerge_start.set(qn("w:val"), "restart")
        tcPr0.append(vMerge_start)
        for ri in range(fabbricato_start + 1, row_idx):
            tc = table.rows[ri].cells[0]._tc
            tcPr = tc.get_or_add_tcPr()
            vMerge_cont = OxmlElement("w:vMerge")
            tcPr.append(vMerge_cont)

    # riga totali
    tot_row = table.rows[row_idx]
    for c in tot_row.cells:
        _set_cell_bg(c, _GRIGIO_INT)
    _cell_text(tot_row.cells[0], "TOTALE", bold=True, size=9, center=True)
    _cell_text(tot_row.cells[1], "", size=8)
    _cell_text(tot_row.cells[2], _fmt(totale_spesa_ante, 2, "€ "), bold=True, size=8, center=True)
    _cell_text(tot_row.cells[3], _fmt(totale_spesa_post,  2, "€ "), bold=True, size=8, center=True)
    _cell_text(tot_row.cells[4], _fmt(totale_risp_ee,    0), bold=True, size=8, center=True)
    _cell_text(tot_row.cells[5], _fmt(totale_risp_euro,  2, "€ "), bold=True, size=8, center=True)
    _cell_text(tot_row.cells[6], _fmt(totale_risp_co2,   1), bold=True, size=8, center=True)


# ─── Tabella investimento ─────────────────────────────────────────────────────

def _tabella_investimento(doc: Document, result: RelampingResult) -> None:
    doc.add_heading("Piano degli investimenti", level=2)

    importo_totale = result.costo_stimato_euro
    n_lampade = sum(r.n_utenze for r in result.righe)
    # Ripartizione standard: 60% corpi, 25% montaggi, 15% sostituzione
    corpi    = importo_totale * 0.60
    montaggi = importo_totale * 0.25
    sost     = importo_totale * 0.15

    # (voce, importo, n_lampade o None)
    voci = [
        ("Importo totale investimento", importo_totale, None),
        ("Sostituzione 1-1",            sost,           None),
        ("Nuovi corpi illuminanti",     corpi,          n_lampade),
        ("Montaggi e sostituzioni",     montaggi,       n_lampade),
    ]

    table = doc.add_table(rows=1 + len(voci), cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    _header_row(table, ["Voce di costo", "N. lampade", "Importo [€]"])

    for i, (voce, importo, n_lamp) in enumerate(voci):
        bg = _VERDE_CHIARO if i % 2 == 0 else _BIANCO
        row = table.rows[i + 1]
        for c in row.cells:
            _set_cell_bg(c, bg)
        bold_flag = (voce == "Importo totale investimento")
        _cell_text(row.cells[0], voce, bold=bold_flag, size=9)
        _cell_text(row.cells[1], str(n_lamp) if n_lamp is not None else "", size=9, center=True)
        _cell_text(row.cells[2], _fmt(importo, 2, "€ "), bold=bold_flag, size=9, center=True)


# ─── Tabelle benchmark VAN ───────────────────────────────────────────────────

def _tabella_benchmark(doc: Document, result: RelampingResult) -> None:
    doc.add_heading("Benchmark economico — Analisi VAN", level=2)

    cb = result.cb
    van_ai = result.van_attualizzato_con_incentivi
    van_as = result.van_attualizzato_senza_incentivi

    # ── Dati in ingresso ──
    doc.add_heading("Dati in ingresso", level=3)
    ingresso = [
        ("Investimento",                  _fmt(van_ai.investimento, 2, "€ ")),
        ("Risparmio annuo [kWh]",         _fmt(van_ai.risparmio_annuo_ee_kwh, 0)),
        ("Risparmio annuo [€/anno]",      _fmt(van_ai.risparmio_annuo_euro, 2, "€ ")),
        ("Durata analisi [anni]",         "8"),
        ("Tasso di attualizzazione",      f"{van_ai.discount_rate * 100:.0f}%"),
        ("Certificati Bianchi [tep/anno]", str(cb.tep_arrotondati)),
        ("Incentivo annuo CB [€/anno]",   _fmt(cb.incentivo_annuo, 2, "€ ")),
        ("Durata incentivi [anni]",       "5"),
    ]
    t_ing = doc.add_table(rows=1 + len(ingresso), cols=2)
    t_ing.style = "Table Grid"
    t_ing.alignment = WD_TABLE_ALIGNMENT.CENTER
    _header_row(t_ing, ["Parametro", "Valore"])
    for i, (param, val) in enumerate(ingresso):
        bg = _VERDE_CHIARO if i % 2 == 0 else _BIANCO
        row = t_ing.rows[i + 1]
        _set_cell_bg(row.cells[0], bg)
        _set_cell_bg(row.cells[1], bg)
        _cell_text(row.cells[0], param, size=9)
        _cell_text(row.cells[1], val, size=9, center=True)

    # ── helper per tabella risultati ──
    indicatori_cols = ["Indicatore", "Att. con incentivi", "Att. senza incentivi"]

    def _fmt_opt(v, decimali=2, suffix=""):
        if v is None:
            return "n.d."
        return _fmt(v, decimali) + suffix

    doc.add_paragraph()
    doc.add_heading("Risultati analisi economica", level=3)
    indicatori = [
        ("VAN [€]",
         _fmt(van_ai.van, 2, "€ "),
         _fmt(van_as.van, 2, "€ ")),
        ("TR semplice [anni]",
         _fmt(van_ai.tr, 1),
         _fmt(van_as.tr, 1)),
        ("DPP [anni]",
         _fmt_opt(van_ai.dpp, 1),
         _fmt_opt(van_as.dpp, 1)),
        ("TIR [%]",
         _fmt_opt(van_ai.tir, 1, "%"),
         _fmt_opt(van_as.tir, 1, "%")),
        ("Indice di profitto",
         _fmt(van_ai.indice_profitto, 3),
         _fmt(van_as.indice_profitto, 3)),
    ]

    t_res = doc.add_table(rows=1 + len(indicatori), cols=3)
    t_res.style = "Table Grid"
    t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    _header_row(t_res, indicatori_cols, font_size=8)
    for i, row_data in enumerate(indicatori):
        bg = _VERDE_CHIARO if i % 2 == 0 else _BIANCO
        row = t_res.rows[i + 1]
        for j, val in enumerate(row_data):
            _set_cell_bg(row.cells[j], bg)
            _cell_text(row.cells[j], val, size=9, center=(j > 0))


# ─── Funzione principale ──────────────────────────────────────────────────────

def costruisci_relazione(results: List[InterventionResult], output_path: str) -> str:
    """
    Crea un documento .docx con i risultati degli interventi.

    Args:
        results: lista di InterventionResult
        output_path: percorso di output .docx

    Returns:
        Percorso assoluto del file generato
    """
    doc = Document()

    # Stile di base
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    doc.add_heading("Relazione Benchmark Energetico", level=0)

    _premessa_generale(doc)

    for result in results:
        if isinstance(result, FVResult):
            _sezione_fotovoltaico(doc, result)
        elif isinstance(result, RelampingResult):
            _sezione_relamping(doc, result)
        else:
            # Sezione generica per futuri interventi
            doc.add_heading(result.nome_intervento.title(), level=1)
            p = doc.add_paragraph()
            p.add_run("Risparmio annuo: ").bold = True
            p.add_run(f"{result.risparmio_annuo_kwh:,.0f} kWh/anno\n")
            p.add_run("Costo stimato: ").bold = True
            p.add_run(f"€ {result.costo_stimato_euro:,.0f}\n")
            p.add_run("Payback: ").bold = True
            p.add_run(f"{result.payback_anni:.1f} anni")

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return str(out.resolve())


def _premessa_generale(doc: Document) -> None:
    """Premessa introduttiva comune agli interventi (testo fisso)."""
    p = doc.add_paragraph()
    p.add_run(
        "L'analisi energetica ha messo in evidenza tre aspetti principali sui quali "
        "focalizzarsi per l'individuazione di interventi di efficientamento energetico:"
    )
    for voce in (
        "sistemi di produzione di energia da fonti rinnovabili;",
        "sostituzione corpi illuminanti;",
        "l'aspetto gestionale dei consumi.",
    ):
        doc.add_paragraph(voce, style="List Bullet")
    doc.add_paragraph(
        "L'installazione di sistemi di produzione di energia da fonti rinnovabili può portare "
        "ad una riduzione rilevante dei costi associati ai consumi di energia elettrica, in quanto "
        "diminuisce la quota di energia prelevata dalla rete. Anche la sostituzione di corpi "
        "illuminanti obsoleti con nuovi a tecnologia LED può risultare rilevante dal punto di vista "
        "di riduzione dei consumi. Nei paragrafi successivi si illustrano nel dettaglio gli "
        "interventi proposti e i miglioramenti che essi potrebbero portare."
    )
    doc.add_paragraph(
        "I costi relativi alla realizzazione degli interventi, riportati nei successivi paragrafi, "
        "sono puramente indicativi sebbene in linea con i costi medi di mercato; solamente una "
        "progettazione accurata e specifica potrà portare ad un valore puntuale dei costi."
    )


# ─── Sezione Fotovoltaico ─────────────────────────────────────────────────────

def _tabella_kv(doc: Document, righe: List[tuple], header: tuple = ("Voce", "Valore"),
                col2_center: bool = True) -> None:
    """Tabella generica chiave/valore a 2 colonne con intestazione verde."""
    table = doc.add_table(rows=1 + len(righe), cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _header_row(table, list(header))
    for i, (k, v) in enumerate(righe):
        bg = _VERDE_CHIARO if i % 2 == 0 else _BIANCO
        row = table.rows[i + 1]
        _set_cell_bg(row.cells[0], bg)
        _set_cell_bg(row.cells[1], bg)
        _cell_text(row.cells[0], str(k), size=9)
        _cell_text(row.cells[1], str(v), size=9, center=col2_center)


def _sezione_fotovoltaico(doc: Document, result: "FVResult") -> None:
    """Genera la sezione completa per l'intervento fotovoltaico."""
    pod_nome = ", ".join(result.pod_selezionati) if result.pod_selezionati else "—"
    n_pannelli_tot = sum(s.n_pannelli for s in result.superfici_input)
    wp_set = sorted({s.potenza_pannello_wp for s in result.superfici_input})
    wp_unit = wp_set[0] if len(wp_set) == 1 else None
    quota_pct = result.quota_autoconsumo * 100.0
    perc_consumo = (result.e_autoconsumata_kwh / result.consumo_pod_kwh * 100.0
                    if result.consumo_pod_kwh > 0 else 0.0)
    fabbricati_txt = ", ".join(result.fabbricati) if result.fabbricati else "<fabbricati>"
    edifici_txt = ", ".join(result.edifici) if result.edifici else "<edifici>"

    doc.add_heading(f"Intervento di installazione impianto fotovoltaico POD {pod_nome}", level=1)

    # ── Intro ──
    intro = doc.add_paragraph()
    intro.add_run(
        "Nel seguente paragrafo si presenta l'intervento di installazione di un impianto solare "
        "fotovoltaico (FTV) per la produzione di energia elettrica da fonte rinnovabile, a servizio "
        f"del POD {pod_nome}. È stato scelto il seguente POD in quanto è tra quelli che presentano "
        "il profilo dei consumi maggiore. "
        f"Sotto il seguente POD sono presenti le utenze relative al {fabbricati_txt}. "
        "Nella presente valutazione dell'intervento, si prevede di considerare solamente gli edifici "
        "le cui coperture permettano un'orientazione ottimale, che garantiscano una maggiore "
        "producibilità, considerando anche eventuali ostacoli ed ombreggiamenti presenti. Nello "
        f"specifico sono state considerate le superfici di copertura dell'edificio {edifici_txt}. "
        "Nella seguente tabella si riportano i dati generali dell'impianto."
    )

    # ── Tabella superfici input ──
    _tabella_superfici_fv(doc, result)

    # ── Paragrafo descrittivo pannelli ──
    wp_txt = f"{_fmt(wp_unit, 0)} Wp" if wp_unit else "tecnologia mista"
    p = doc.add_paragraph()
    p.add_run(
        "Per l'installazione, sono stati scelti dei pannelli in silicio monocristallino, "
        f"caratterizzati da una potenza di picco pari a {wp_txt}. L'impianto costituito da nr. "
        f"{n_pannelli_tot} unità, in grado di produrre annualmente {_fmt(result.e_prodotta_kwh, 0)} "
        f"kWh/anno; la quota destinata all'autoconsumo è pari a "
        f"{_fmt(result.e_autoconsumata_kwh, 0)} kWh/anno, ovvero circa il "
        f"{_fmt(perc_consumo, 1)}% dei consumi del POD."
    )

    # ── Tabella campo FV ──
    _tabella_kv(doc, [
        ("Potenza unitaria [Wp]", _fmt(wp_unit, 0) if wp_unit else "varie"),
        ("Potenza installata [kWp]", _fmt(result.potenza_totale_kwp, 2)),
        ("Producibilità [kWh/anno]", _fmt(result.e_prodotta_kwh, 0)),
    ], header=("Pannello FV monocristallino", "Valore"))

    # ── Grafico mensile ──
    doc.add_heading("Confronto mensile produzione/consumi", level=2)
    doc.add_paragraph(
        f"Nel grafico sottostante si confronta la produzione dell'impianto fotovoltaico con i "
        f"consumi mensili del POD {pod_nome}, dove si nota il sovradimensionamento dell'impianto."
    )
    buf = grafico_mensile_prod_vs_cons(result.produzione_oraria, result.consumo_orario)
    doc.add_picture(buf, width=Cm(15))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── Grafici giornate tipiche (22 luglio estiva, 22 gennaio invernale) ──
    doc.add_heading("Curve di carico giornaliere tipiche", level=2)
    buf = grafico_giornata_tipica(result.produzione_oraria, result.consumo_orario,
                                  mese=7, giorno=22, titolo="Giornata tipica estiva — 22 luglio")
    doc.add_picture(buf, width=Cm(15))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    buf = grafico_giornata_tipica(result.produzione_oraria, result.consumo_orario,
                                  mese=1, giorno=22, titolo="Giornata tipica invernale — 22 gennaio")
    doc.add_picture(buf, width=Cm(15))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── Confronto stagionale (estate giu-lug-ago, inverno dic-gen-feb) ──
    doc.add_heading("Confronto stagionale", level=2)
    buf = grafico_stagionale(result.produzione_oraria, result.consumo_orario,
                             mesi=[6, 7, 8], titolo="Stagione estiva (giugno–agosto)")
    doc.add_picture(buf, width=Cm(15))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    buf = grafico_stagionale(result.produzione_oraria, result.consumo_orario,
                             mesi=[12, 1, 2], titolo="Stagione invernale (dicembre–febbraio)")
    doc.add_picture(buf, width=Cm(15))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── Valutazione economica ──
    _sezione_economica_fv(doc, result, pod_nome)


def _tabella_superfici_fv(doc: Document, result: "FVResult") -> None:
    cols = ["Descrizione", "Inclinazione [°]", "Azimuth [°]", "Nr. pannelli", "kWp installati"]
    sups = result.superfici_input
    table = doc.add_table(rows=1 + len(sups) + 1, cols=len(cols))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _header_row(table, cols)

    for i, s in enumerate(sups):
        bg = _VERDE_CHIARO if i % 2 == 0 else _BIANCO
        row = table.rows[i + 1]
        for c in row.cells:
            _set_cell_bg(c, bg)
        _cell_text(row.cells[0], s.id, size=8)
        _cell_text(row.cells[1], _fmt(s.slope, 0), size=8, center=True)
        _cell_text(row.cells[2], _fmt(s.azimuth, 0), size=8, center=True)
        _cell_text(row.cells[3], str(s.n_pannelli), size=8, center=True)
        _cell_text(row.cells[4], _fmt(s.peak_power_kwp, 2), size=8, center=True)

    # riga totale
    tot = table.rows[-1]
    for c in tot.cells:
        _set_cell_bg(c, _GRIGIO_INT)
    _cell_text(tot.cells[0], "TOTALE", bold=True, size=8, center=True)
    _cell_text(tot.cells[1], "", size=8)
    _cell_text(tot.cells[2], "", size=8)
    _cell_text(tot.cells[3], str(sum(s.n_pannelli for s in sups)), bold=True, size=8, center=True)
    _cell_text(tot.cells[4], _fmt(result.potenza_totale_kwp, 2), bold=True, size=8, center=True)


def _sezione_economica_fv(doc: Document, result: "FVResult", pod_nome: str) -> None:
    doc.add_heading("Valutazione economica dell'intervento", level=2)

    consumo_pod = result.consumo_pod_kwh
    e_prel = result.e_prelevata_kwh
    fattore_co2 = 294.784 / 1000
    emiss_pre = consumo_pod * fattore_co2
    emiss_post = e_prel * fattore_co2

    va = result.van_attualizzato
    vs = result.van_semplice

    # €/kWp riferito all'investimento totale (impianto + progettazione), come nel riferimento
    inv_per_kwp = (result.investimento_totale_euro / result.potenza_totale_kwp
                   if result.potenza_totale_kwp > 0 else 0.0)

    doc.add_paragraph(
        "Si riportano di seguito i dati relativi all'installazione dell'impianto di produzione "
        "fotovoltaica e l'analisi economica, con il calcolo del tempo di ritorno semplice ed "
        "attualizzato e dei principali indicatori economici. Il costo di investimento iniziale è "
        "stato valutato secondo prezzi indicati dal prezzario regionale, per un totale di "
        f"{_fmt(inv_per_kwp, 0)} €/kWp tenendo conto anche dei costi legati alla "
        f"progettazione. Per la stima dei risparmi è stato considerato un costo medio dell'energia "
        f"pari a {_fmt(result.prezzo_kwh, 3)} €/kWh. Si è stimato un autoconsumo del "
        f"{_fmt(result.quota_autoconsumo * 100, 1)}%, analizzando i consumi mensili e le curve di "
        "carico a disposizione."
    )
    doc.add_paragraph(
        "I benefici annui correlati all'intervento sono dovuti sia al mancato acquisto da rete, sia "
        "all'immissione in rete, remunerata tramite il meccanismo del Ritiro Dedicato, che valorizza "
        "l'immissione ai prezzi minimi garantiti stabiliti da Arera o a prezzo di mercato secondo il "
        f"PUN (Prezzo Unico Nazionale). Si prende come riferimento la media del PUN monorario "
        f"dell'anno {result.anno_pvgis} pari a {_fmt(result.pun_euro_kwh, 3)} €/kWh."
    )

    # ── Tabella installazione (energetica + ambientale) ──
    _tabella_kv(doc, [
        ("Potenza installata [kWp]", _fmt(result.potenza_totale_kwp, 2)),
        ("Producibilità specifica [kWh/kWp]", _fmt(result.ore_equivalenti_impianto, 0)),
        ("Producibilità conseguibile [kWh/anno]", _fmt(result.e_prodotta_kwh, 0)),
        ("Percentuale utilizzabile in autoconsumo [%]", _fmt(result.quota_autoconsumo * 100, 1)),
        ("Energia autoconsumata [kWh/anno]", _fmt(result.e_autoconsumata_kwh, 0)),
        ("Consumo EE PRE INTERVENTO [kWh/anno]", _fmt(consumo_pod, 0)),
        ("Consumo EE POST INTERVENTO [kWh/anno]", _fmt(e_prel, 0)),
        ("Mancato prelievo di EE da rete [kWh/anno]", _fmt(result.e_autoconsumata_kwh, 0)),
        ("Mancato prelievo di EE da rete [TEP]", _fmt(result.tep_risparmiati, 3)),
        ("Emissioni di CO₂ PRE INTERVENTO [kg/anno]", _fmt(emiss_pre, 0)),
        ("Emissioni di CO₂ POST INTERVENTO [kg/anno]", _fmt(emiss_post, 0)),
        ("Risparmio CO₂ conseguibile [kg/anno]", _fmt(result.co2_evitata_kg, 0)),
    ], header=("Installazione impianto fotovoltaico", "Valore"))

    # ── Grafico costi ante/post + immissione ──
    doc.add_heading("Confronto economico ante/post intervento", level=3)
    doc.add_paragraph(
        "Si riporta di seguito la spesa economica ante e post intervento, considerando anche "
        "l'immissione in rete. Si nota un buon guadagno associato all'immissione, dovuto "
        "all'installazione di un impianto sovradimensionato rispetto alle esigenze del POD."
    )
    costo_ante = consumo_pod * result.prezzo_kwh
    costo_post = e_prel * result.prezzo_kwh
    buf = grafico_costi_fv(costo_ante, costo_post, result.ricavo_immissione_euro)
    doc.add_picture(buf, width=Cm(12))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── Tabella investimento ──
    doc.add_heading("Stima dell'investimento iniziale", level=3)
    _tabella_kv(doc, [
        ("Costo impianto", _fmt(result.costo_impianto_euro, 2, "€ ")),
        ("Progettazione (15%)", _fmt(result.progettazione_euro, 2, "€ ")),
        ("Investimento iniziale totale", _fmt(result.investimento_totale_euro, 2, "€ ")),
    ], header=("Voce di costo", "Importo"))

    # ── Tabella valutazione economica senza incentivi ──
    doc.add_heading("Valutazione economica", level=3)

    # TR attualizzato: se il VAN resta negativo entro la vita utile, non rientra → "> N anni"
    dpp_txt = f"{_fmt(va.dpp, 2)}" if va.dpp is not None else f"> {result.vita_utile_anni} anni"
    tir_txt = "n.d." if va.tir is None else _fmt(va.tir, 2) + "%"

    _tabella_kv(doc, [
        ("Investimento iniziale [€]", _fmt(result.investimento_totale_euro, 0)),
        ("Risparmio annuale [€]", _fmt(result.risparmio_acquisto_euro, 0)),
        ("Tasso di interesse [%]", _fmt(va.discount_rate * 100, 1) + "%"),
        ("Vita utile investimento [anni]", str(result.vita_utile_anni)),
        ("Spese aggiuntive annuali (manutenzione ordinaria e gestione) [€]",
         _fmt(result.manutenzione_annua_euro, 0)),
        ("Guadagno da immissione in rete [€]", _fmt(result.ricavo_immissione_euro, 0)),
        ("Valore attuale netto (VAN) [€]", _fmt(va.van, 0)),
        ("Tempo di ritorno [anni]", _fmt(vs.tr, 2)),
        ("Tempo di ritorno attualizzato [anni]", dpp_txt),
        ("Tasso interno di rendimento (TIR) [%]", tir_txt),
        ("Indice di profitto [p.u.]", _fmt(va.indice_profitto, 2)),
    ], header=("Valutazione economica senza incentivi", "Valore"))

    # ── Grafico flussi di cassa ──
    doc.add_heading("Analisi dei flussi di cassa", level=3)
    buf = grafico_flussi_cassa_fv(
        fc_attualizzato=va.flussi_cassa,
        fc_semplice=vs.flussi_cassa,
        investimento=result.investimento_totale_euro,
    )
    doc.add_picture(buf, width=Cm(15))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def _sezione_relamping(doc: Document, result: RelampingResult) -> None:
    """Genera la sezione completa per l'intervento di relamping."""

    doc.add_heading("Intervento di Relamping", level=1)

    # ── 1. Paragrafo introduttivo ──
    intro = doc.add_paragraph()
    intro.add_run(
        "L'intervento si propone di valutare i risparmi ottenibili dalla sostituzione dei corpi "
        "illuminanti esistenti con tecnologia LED ad alta efficienza. "
        "La sostituzione dei corpi illuminanti con analoghi apparecchi a LED permette di ottenere "
        "una riduzione dei consumi elettrici anche superiore al 50%, con un significativo "
        "miglioramento della qualità della luce e una drastica riduzione delle operazioni di "
        "manutenzione. L'analisi tiene conto delle caratteristiche specifiche di ogni tipologia "
        "di lampada presente nell'edificio, applicando fattori di conversione differenziati per "
        "categoria tecnologica."
    )

    # ── 2. Tabella 1 — Valutazione relamping ──
    _tabella_valutazione_relamping(doc, result)

    # ── 3. Tabella 2 — Valutazione energetica e economica ──
    _tabella_economica(doc, result)

    # ── 4. Grafico spesa ante/post ──
    doc.add_heading("Confronto spesa energetica", level=2)
    totale_ante = sum(r.spesa_ante_euro for r in result.righe)
    totale_post = sum(r.spesa_post_euro for r in result.righe)
    buf_spesa = grafico_spesa_ante_post(totale_ante, totale_post)
    doc.add_picture(buf_spesa, width=Cm(14))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── 5. Tabella investimento ──
    _tabella_investimento(doc, result)

    # ── 6. Certificati Bianchi ──
    doc.add_heading("Certificati Bianchi", level=2)
    cb = result.cb
    cb_para = doc.add_paragraph()
    cb_para.add_run(
        "L'intervento di relamping può beneficiare del meccanismo dei Certificati Bianchi "
        "(Titoli di Efficienza Energetica — TEE), riconosciuti dal Gestore dei Servizi "
        "Energetici (GSE) per interventi di efficienza energetica nel settore dell'illuminazione. "
        "Sulla base dei calcoli effettuati, il risparmio stimato è pari a "
    )
    cb_para.add_run(f"{cb.tep_risparmiati:.3f} tep/anno").bold = True
    cb_para.add_run(", corrispondente a ")
    cb_para.add_run(f"{cb.tep_arrotondati} tep/anno").bold = True
    cb_para.add_run(
        " (arrotondati per difetto come da normativa). "
        "Considerando un valore di mercato pari a 250 €/tep, l'incentivo annuo stimato è di "
    )
    cb_para.add_run(f"€ {cb.incentivo_annuo:,.2f}/anno").bold = True
    cb_para.add_run(" per i primi 5 anni dall'intervento.")

    # ── 7. Tabelle benchmark ──
    _tabella_benchmark(doc, result)

    # ── 8. Grafico flussi di cassa ──
    doc.add_heading("Analisi dei flussi di cassa", level=2)
    buf_fc = grafico_flussi_cassa(
        fc_att_con=result.van_attualizzato_con_incentivi.flussi_cassa,
        fc_att_senza=result.van_attualizzato_senza_incentivi.flussi_cassa,
        fc_sempl_con=result.van_semplice_con_incentivi.flussi_cassa,
        fc_sempl_senza=result.van_semplice_senza_incentivi.flussi_cassa,
        investimento=result.van_attualizzato_con_incentivi.investimento,
    )
    doc.add_picture(buf_fc, width=Cm(14))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    if result.note:
        doc.add_heading("Note", level=2)
        for nota in result.note:
            doc.add_paragraph(nota, style="List Bullet")
