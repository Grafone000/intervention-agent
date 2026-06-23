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
from reporting.charts import grafico_spesa_ante_post, grafico_flussi_cassa


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

    for result in results:
        if isinstance(result, RelampingResult):
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
