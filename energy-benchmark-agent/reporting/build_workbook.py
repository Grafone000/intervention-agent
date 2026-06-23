"""
Genera un workbook .xlsx con FORMULE LIVE che documenta tutti i calcoli
del relamping, ricalcando la struttura dell'Excel di riferimento
(Sostituzione_Lampade_Salerno.xlsx).

Obiettivo: verificabilità. Ogni valore derivato è scritto come formula
Excel (=...) che fa riferimento alle celle di input, così l'utente può
aprire il file e controllare la catena di calcolo cella per cella.

Fogli prodotti:
  - FATTORI CONVERSIONE : costanti (tep, CO2, costi unitari)
  - Calcoli             : una riga per lampada, formule G/J/K/M/N/O/P/Q/R/S
  - CB                  : certificati bianchi (perimetro = tutte le righe)
  - Dati generali       : totali energetici/economici e indicatori
  - TIR                 : flussi per il calcolo IRR (con/senza incentivi)
  - VAN attualizzato con incentivi
  - VAN attualizzato no incentivi
  - VAN semplice con incentivi
  - VAN semplice no incentivi

I valori di INPUT (n. lampade, potenze ante/post, costo lampada, h eq.)
provengono dal RelampingResult; tutto il resto è formula.
"""

from __future__ import annotations
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from interventions.relamping import RelampingResult

# ── nomi fogli (coerenti tra loro, senza spazi finali) ───────────────────────
FATT   = "FATTORI CONVERSIONE"
CALC   = "Calcoli"
CB     = "CB"
DATI   = "Dati generali"
TIR    = "TIR"
VAN_AC = "VAN attualizzato con incentivi"
VAN_AN = "VAN attualizzato no incentivi"
VAN_SC = "VAN semplice con incentivi"
VAN_SN = "VAN semplice no incentivi"

# ── stili ────────────────────────────────────────────────────────────────────
_F_TITOLO  = Font(bold=True, size=12, color="FFFFFF")
_F_HEADER  = Font(bold=True, size=10, color="FFFFFF")
_F_BOLD    = Font(bold=True, size=10)
_FILL_HEAD = PatternFill("solid", fgColor="1F7535")
_FILL_TIT  = PatternFill("solid", fgColor="375623")
_FILL_INP  = PatternFill("solid", fgColor="FFF2CC")   # celle di input (giallo)
_FILL_TOT  = PatternFill("solid", fgColor="E2EFDA")
_CENTER    = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT      = Alignment(horizontal="left", vertical="center")
_thin = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

_FMT_EUR = '#,##0.00 "€"'
_FMT_KWH = '#,##0.00'
_FMT_KW  = '#,##0.000'
_FMT_PCT = '0.00%'


def _hdr(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = _F_HEADER
        c.fill = _FILL_HEAD
        c.alignment = _CENTER
        c.border = _BORDER


# ─────────────────────────────────────────────────────────────────────────────
# FATTORI CONVERSIONE
# ─────────────────────────────────────────────────────────────────────────────
def _sheet_fattori(wb: Workbook) -> None:
    ws = wb.create_sheet(FATT)
    ws["A1"] = "Fattori di conversione in energia primaria"
    ws["B1"] = "u.m."
    ws["C1"] = "Valore"
    for cell in ("A1", "B1", "C1"):
        ws[cell].font = _F_BOLD

    # tep/unità — C2 è referenziato come fattore tep elettricità
    righe_tep = [
        ("Energia elettrica", "tep/kWh", 0.000187),
        ("Gas Naturale", "tep/Sm3 (tep/Smc)", 0.000836),
        ("GPL", "tep/kg", 0.0011),
        ("Gasolio", "tep/l", 0.00086),
        ("Benzina", "tep/l", 0.000765),
    ]
    for i, (nome, um, val) in enumerate(righe_tep):
        r = 2 + i
        ws.cell(row=r, column=1, value=nome)
        ws.cell(row=r, column=2, value=um)
        ws.cell(row=r, column=3, value=val)

    # Fattori di emissione CO2 — C15 elettricità
    ws["A14"] = "Fattori di emissione anidride carbonica"
    ws["B14"] = "u.m."
    ws["C14"] = "Valore"
    for cell in ("A14", "B14", "C14"):
        ws[cell].font = _F_BOLD
    righe_co2 = [
        ("Energia elettrica", "kg/MWh", 294.784),
        ("Gas naturale", "kg/Sm3", 2.004),
        ("GPL", "kg/kg", 3.026),
        ("Gasolio", "kg/l", 2.65),
        ("Benzina", "kg/l", 2.34),
    ]
    for i, (nome, um, val) in enumerate(righe_co2):
        r = 15 + i
        ws.cell(row=r, column=1, value=nome)
        ws.cell(row=r, column=2, value=um)
        ws.cell(row=r, column=3, value=val)

    # Costi unitari — C24 elettricità €/kWh
    ws["A23"] = "Costi unitari"
    ws["B23"] = "u.m."
    ws["C23"] = "Valore"
    for cell in ("A23", "B23", "C23"):
        ws[cell].font = _F_BOLD
    righe_costi = [
        ("Energia elettrica", "€/kWh", 0.21),
        ("Gas naturale", "€/Smc", 0.745),
        ("GPL", "€/lt", 1.159),
        ("Gasolio", "€/lt", 1.08),
    ]
    for i, (nome, um, val) in enumerate(righe_costi):
        r = 24 + i
        ws.cell(row=r, column=1, value=nome)
        ws.cell(row=r, column=2, value=um)
        ws.cell(row=r, column=3, value=val)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14


# ─────────────────────────────────────────────────────────────────────────────
# Calcoli
# ─────────────────────────────────────────────────────────────────────────────
def _sheet_calcoli(wb: Workbook, result: RelampingResult) -> int:
    """Ritorna l'indice dell'ultima riga dati (per i SUM degli altri fogli)."""
    ws = wb.create_sheet(CALC)
    # Colonne A(1)..V(22)
    cols = [
        "Edificio/zona",            # A  1
        "Competenza",               # B  2
        "POD",                      # C  3
        "Tipo di illuminazione",    # D  4
        "N. lampade",               # E  5  INPUT
        "P lampada ante [kW]",      # F  6  INPUT
        "P totale ante [kW]",       # G  7  =F*E
        "P lampada post [kW]",      # H  8  INPUT
        "Costo lampada [€]",        # I  9  INPUT
        "Costo totale [€]",         # J 10  =E*I
        "P totale post [kW]",       # K 11  =E*H
        "h eq.",                    # L 12  INPUT
        "F. contemporaneità",       # M 13  INPUT
        "Consumo POST [kWh]",       # N 14  =K*L*M
        "Consumo ANTE [kWh]",       # O 15  =G*L*M
        "Spesa ANTE [€]",           # P 16  =O*C24
        "Spesa POST [€]",           # Q 17  =N*C24
        "Emissioni ANTE [kg/anno]", # R 18  =O*C15/1000
        "Emissioni POST [kg/anno]", # S 19  =N*C15/1000
        "Risparmio EE [kWh]",       # T 20  =O-N
        "Risparmio € [€]",          # U 21  =P-Q
        "Risparmio CO2 [kg]",       # V 22  =R-S
    ]
    _hdr(ws, 1, cols)

    for idx, riga in enumerate(result.righe):
        r = 2 + idx

        # celle di sola lettura (non input)
        ws.cell(row=r, column=1, value=riga.zona)               # A edificio/zona
        ws.cell(row=r, column=2, value=riga.zona)               # B competenza (stesso campo)
        ws.cell(row=r, column=3, value=riga.pod)                # C POD
        ws.cell(row=r, column=4, value=riga.tipologia_originale)# D tipo

        # celle di input (giallo)
        ce = ws.cell(row=r, column=5,  value=riga.n_utenze)              # E
        cf = ws.cell(row=r, column=6,  value=riga.p_lampada_ante_kw)     # F
        ch = ws.cell(row=r, column=8,  value=riga.p_lampada_post_kw)     # H
        ci = ws.cell(row=r, column=9,  value=riga.costo_lampada_euro)    # I
        cl = ws.cell(row=r, column=12, value=riga.h_equivalenti)         # L
        cm = ws.cell(row=r, column=13, value=riga.fatt_contemporaneita)  # M
        for cc in (ce, cf, ch, ci, cl, cm):
            cc.fill = _FILL_INP

        # formule
        ws.cell(row=r, column=7,  value=f"=F{r}*E{r}")                          # G P tot ante
        ws.cell(row=r, column=10, value=f"=E{r}*I{r}")                          # J costo tot
        ws.cell(row=r, column=11, value=f"=E{r}*H{r}")                          # K P tot post
        ws.cell(row=r, column=14, value=f"=K{r}*L{r}*M{r}")                     # N consumo post
        ws.cell(row=r, column=15, value=f"=G{r}*L{r}*M{r}")                     # O consumo ante
        ws.cell(row=r, column=16, value=f"=O{r}*'{FATT}'!$C$24")                # P spesa ante
        ws.cell(row=r, column=17, value=f"=N{r}*'{FATT}'!$C$24")                # Q spesa post
        ws.cell(row=r, column=18, value=f"=O{r}*'{FATT}'!$C$15/1000")           # R emiss ante
        ws.cell(row=r, column=19, value=f"=N{r}*'{FATT}'!$C$15/1000")           # S emiss post
        ws.cell(row=r, column=20, value=f"=O{r}-N{r}")                           # T risp EE
        ws.cell(row=r, column=21, value=f"=P{r}-Q{r}")                           # U risp €
        ws.cell(row=r, column=22, value=f"=R{r}-S{r}")                           # V risp CO2

    last = 1 + len(result.righe)

    # blocco sommario in X/Y (col 24/25)
    somm = [
        ("Consumo EE Ante [kWh]", f"=SUM(O2:O{last})"),
        ("Consumo EE Post [kWh]", f"=SUM(N2:N{last})"),
        ("Capex [€]",             f"=SUM(J2:J{last})"),
        ("Spesa Ante",            f"=SUM(P2:P{last})"),
        ("Spesa Post",            f"=SUM(Q2:Q{last})"),
    ]
    for k, (lab, formula) in enumerate(somm):
        r = 4 + k
        cl = ws.cell(row=r, column=24, value=lab)
        cl.font = _F_BOLD
        cv = ws.cell(row=r, column=25, value=formula)
        cv.fill = _FILL_TOT
        cv.font = _F_BOLD

    # larghezze colonne A-V
    widths = [16, 16, 14, 22, 10, 14, 14, 14, 12, 12, 14, 9, 14, 14, 14, 12, 12, 16, 16, 14, 12, 14]
    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.column_dimensions["X"].width = 22
    ws.column_dimensions["Y"].width = 16
    ws.freeze_panes = "E2"
    return last


# ─────────────────────────────────────────────────────────────────────────────
# CB — Certificati Bianchi (perimetro = tutte le righe)
# ─────────────────────────────────────────────────────────────────────────────
def _sheet_cb(wb: Workbook, last: int) -> None:
    ws = wb.create_sheet(CB)
    ws["B2"] = "CERTIFICATI BIANCHI"
    ws["B2"].font = _F_BOLD
    voci = [
        ("Pbaseline", "[kW]",   f"=SUM({CALC}!G2:G{last})"),
        ("Epost",     "[kWh]",  f"=SUM({CALC}!N2:N{last})"),
        ("Ppost",     "[kW]",   f"=SUM({CALC}!K2:K{last})"),
        ("hpost",     "[h]",    "=D4/D5"),
        ("REA",       "[Tep]",  f"=(D3*D6-D4)*'{FATT}'!C2"),
    ]
    for k, (lab, um, formula) in enumerate(voci):
        r = 3 + k
        ws.cell(row=r, column=2, value=lab).font = _F_BOLD
        ws.cell(row=r, column=3, value=um)
        ws.cell(row=r, column=4, value=formula)
    ws.cell(row=7, column=5, value="=ROUND(D7,0)")   # E7 REA arrotondato
    ws.cell(row=8, column=2, value="CB").font = _F_BOLD
    ws.cell(row=8, column=3, value="[€/tep]")
    ws.cell(row=8, column=4, value=250)
    ws.cell(row=9, column=2, value="Incentivo annuo").font = _F_BOLD
    ws.cell(row=9, column=3, value="[€]")
    ws.cell(row=9, column=4, value="=D8*E7")
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 12


# ─────────────────────────────────────────────────────────────────────────────
# Dati generali
# ─────────────────────────────────────────────────────────────────────────────
def _sheet_dati(wb: Workbook, result: RelampingResult, last: int) -> None:
    ws = wb.create_sheet(DATI)

    def tit(r, txt):
        c = ws.cell(row=r, column=2, value=txt)
        c.font = _F_TITOLO
        c.fill = _FILL_TIT
        ws.cell(row=r, column=3).fill = _FILL_TIT
        ws.cell(row=r, column=4).fill = _FILL_TIT

    def riga(r, lab, um, val, fmt=None):
        ws.cell(row=r, column=2, value=lab)
        ws.cell(row=r, column=3, value=um)
        c = ws.cell(row=r, column=4, value=val)
        if fmt:
            c.number_format = fmt

    tit(3, "SOSTITUZIONE IMPIANTO ILLUMINAZIONE")
    riga(4,  "Potenza installata", "[kW]", f"=SUM({CALC}!G2:G{last})", _FMT_KW)
    riga(5,  "Consumo EE PRE INTERVENTO", "[kWh]", f"={CALC}!Y4", _FMT_KWH)
    riga(6,  "Consumo EE POST INTERVENTO", "[kWh]", f"={CALC}!Y5", _FMT_KWH)
    riga(7,  "Risparmio EE conseguibile", "[kWh]", "=D5-D6", _FMT_KWH)
    riga(8,  "Risparmio energia primaria", "[Tep]", f"=D7*'{FATT}'!C2", _FMT_KWH)
    riga(9,  "Emissioni di CO2 PRE INTERVENTO", "[kg/anno]", f"=D5*'{FATT}'!$C$15/1000", _FMT_KWH)
    riga(10, "Emissioni di CO2 POST INTERVENTO", "[kg/anno]", f"=D6*'{FATT}'!$C$15/1000", _FMT_KWH)
    riga(11, "Risparmio CO2 conseguibile", "[kg/anno]", "=D9-D10", _FMT_KWH)

    tit(12, "VALUTAZIONE ECONOMICA")
    riga(13, "Investimento iniziale", "[€]", f"={CALC}!Y6", _FMT_EUR)
    riga(14, "Risparmio annuale", "[€]", f"=D7*'{FATT}'!C24", _FMT_EUR)
    riga(15, "Tasso di interesse", "[%]", 0.06, _FMT_PCT)
    riga(16, "Ritorno incentivi annuale", "[€]", f"={CB}!D9", _FMT_EUR)
    riga(17, "Spese aggiuntive", "[€]", 0, _FMT_EUR)
    riga(18, "Frequenza spese aggiuntive", "[anni]", 1)
    riga(19, "Vita utile investimento", "[anni]", 8)
    riga(20, "Anni incentivo", "[anni]", 5)

    tit(21, "RISULTATI ECONOMICI CON INCENTIVI")
    riga(22, "Valore attuale netto (VAN)", "[€]", f"='{VAN_AC}'!G13", _FMT_EUR)
    riga(23, "Tempo di ritorno", "[anni]", "=(D13-D16*D20)/D14")
    riga(24, "Tempo di ritorno attualizzato", "[anni]",
         result.van_attualizzato_con_incentivi.dpp)
    riga(25, "Tasso interno di rendimento (TIR)", "[%]", f"={TIR}!E27", _FMT_PCT)
    riga(26, "Indice di profitto", "[p.u.]", "=D22/D13")

    tit(27, "RISULTATI ECONOMICI SENZA INCENTIVI")
    riga(28, "Valore attuale netto (VAN)", "[€]", f"='{VAN_AN}'!G13", _FMT_EUR)
    riga(29, "Tempo di ritorno", "[anni]", "=D13/D14")
    riga(30, "Tempo di ritorno attualizzato", "[anni]",
         result.van_attualizzato_senza_incentivi.dpp)
    riga(31, "Tasso interno di rendimento (TIR)", "[%]", f"={TIR}!H26", _FMT_PCT)
    riga(32, "Indice di profitto", "[p.u.]", "=D28/D13")

    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16


# ─────────────────────────────────────────────────────────────────────────────
# VAN (parametrico)
# ─────────────────────────────────────────────────────────────────────────────
def _sheet_van(wb: Workbook, nome: str, discount_rate: float,
               con_incentivi: bool, anni_orizzonte: int = 20,
               vita: int = 8, durata_inc: int = 5) -> None:
    ws = wb.create_sheet(nome)
    ws["B2"] = "Year"
    ws["E2"] = "Discount rate"
    ws["F2"] = discount_rate
    ws["G2"] = "RELAMPING"
    ws["C3"] = "CC1="
    ws["D3"] = f"={CALC}!Y6"
    ws["C4"] = "Net Annual Saving - Risparmio annuo"
    ws["G4"] = "Cash Flow (VAN)"

    for n in range(0, anni_orizzonte + 1):
        r = 5 + n
        ws.cell(row=r, column=2, value=n)                       # B year
        ws.cell(row=r, column=5, value=f"=1/(1+$F$2)^B{r}")     # E fattore att.
        if n == 0:
            ws.cell(row=r, column=9, value=f"=C{r}*E{r}")       # I = 0
            ws.cell(row=r, column=7, value="=-$D$3")            # G = -Inv
        else:
            if con_incentivi and n <= durata_inc:
                ws.cell(row=r, column=3, value=f"='{DATI}'!$D$14+{CB}!$D$9")
            else:
                ws.cell(row=r, column=3, value=f"='{DATI}'!$D$14")
            ws.cell(row=r, column=9, value=f"=C{r}*E{r}")            # I = C*E attualizzato
            ws.cell(row=r, column=7, value=f"=-$D$3+SUM($I$5:I{r})")  # G cumulato

    # evidenzia la riga del VAN (anno = vita)
    r_van = 5 + vita
    for col in (2, 3, 5, 7, 9):
        ws.cell(row=r_van, column=col).font = _F_BOLD
        ws.cell(row=r_van, column=col).fill = _FILL_TOT

    ws.cell(row=r_van + 1, column=6, value=f"VAN a {vita} anni →").font = _F_BOLD
    ws.cell(row=r_van + 1, column=7, value=f"=G{r_van}").font = _F_BOLD

    for col, w in {"B": 8, "C": 16, "D": 4, "E": 14, "F": 10, "G": 18, "I": 16}.items():
        ws.column_dimensions[col].width = w


# ─────────────────────────────────────────────────────────────────────────────
# TIR
# ─────────────────────────────────────────────────────────────────────────────
def _sheet_tir(wb: Workbook, anni_orizzonte: int = 20,
               vita: int = 8, durata_inc: int = 5) -> None:
    ws = wb.create_sheet(TIR)
    ws["C3"] = "INTERVENTO DI SOSTITUZIONE IMPIANTO D'ILLUMINAZIONE"
    ws["C3"].font = _F_BOLD
    _hdr(ws, 4, ["VITA DEL COMPONENTE", "USCITE (INVESTIMENTO)",
                 "RISPARMIO ENERGETICO", "CB", "TOTALE",
                 "TOTALE SENZA INCENTIVI"], start_col=3)

    # anno 0
    ws["C5"] = 0
    ws["D5"] = f"=-{CALC}!Y6"
    ws["E5"] = "-"
    ws["F5"] = "-"
    ws["G5"] = "=D5"
    ws["H5"] = "=D5"

    for n in range(1, anni_orizzonte + 1):
        r = 5 + n
        ws.cell(row=r, column=3, value=n)
        ws.cell(row=r, column=4, value="-")
        ws.cell(row=r, column=5, value=f"='{DATI}'!$D$14")     # E risparmio
        if n <= durata_inc:
            ws.cell(row=r, column=6, value=f"={CB}!$D$9")       # F CB
            ws.cell(row=r, column=7, value=f"=E{r}+F{r}")        # G totale
        else:
            ws.cell(row=r, column=6, value="-")
            ws.cell(row=r, column=7, value=f"=E{r}")
        ws.cell(row=r, column=8, value=f"=E{r}")                 # H senza incentivi

    r_irr = 5 + anni_orizzonte + 1   # riga sotto l'ultimo anno
    r_van = 5 + vita                 # IRR calcolato sull'orizzonte di vita
    ws.cell(row=r_irr, column=3, value="TIR").font = _F_BOLD
    ws.cell(row=r_irr, column=4, value=f"=IRR(G5:G{r_van})")     # D26 con incentivi
    ws.cell(row=r_irr, column=8, value=f"=IRR(H5:H{r_van})")     # H26 senza incentivi
    ws.cell(row=r_irr + 1, column=5, value=f"=D{r_irr}")         # E27 = D26
    for c in ("D", "H"):
        ws[f"{c}{r_irr}"].number_format = _FMT_PCT
        ws[f"{c}{r_irr}"].font = _F_BOLD
    ws[f"E{r_irr + 1}"].number_format = _FMT_PCT

    for col, w in {"C": 20, "D": 18, "E": 18, "F": 12, "G": 16, "H": 20}.items():
        ws.column_dimensions[col].width = w


# ─────────────────────────────────────────────────────────────────────────────
# entrypoint
# ─────────────────────────────────────────────────────────────────────────────
def costruisci_workbook(result: RelampingResult, output_path: str) -> str:
    """
    Genera il workbook .xlsx con formule live per l'intervento di relamping.

    Args:
        result: RelampingResult da interventions.relamping.calcola()
        output_path: percorso .xlsx di output

    Returns:
        Percorso assoluto del file generato
    """
    wb = Workbook()
    wb.remove(wb.active)   # rimuove il foglio di default

    _sheet_fattori(wb)
    last = _sheet_calcoli(wb, result)
    _sheet_cb(wb, last)
    _sheet_dati(wb, result, last)
    _sheet_tir(wb)
    _sheet_van(wb, VAN_AC, discount_rate=0.06, con_incentivi=True)
    _sheet_van(wb, VAN_AN, discount_rate=0.06, con_incentivi=False)
    _sheet_van(wb, VAN_SC, discount_rate=0.0,  con_incentivi=True)
    _sheet_van(wb, VAN_SN, discount_rate=0.0,  con_incentivi=False)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out))
    return str(out.resolve())
