"""
build_workbook.py
-----------------
Step 6 + output - writes output/Oracle_GL_Migration_Simulator.xlsx from the
CSVs produced by the earlier steps, adds the fit-gap analysis, and wires every
KPI, flag, control check and reconciliation line as a live Excel formula.

Run order: generate_legacy_data.py -> mapping.py -> transform.py ->
reconcile.py -> build_workbook.py
"""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from mapping import LEGACY_FS_TO_GROUP, LOB, NATURAL_ACCOUNT

ROOT = Path(__file__).resolve().parents[1]
LEGACY = ROOT / "data" / "legacy"
CLOUD = ROOT / "data" / "cloud"
OUT = ROOT / "output" / "Oracle_GL_Migration_Simulator.xlsx"

COMPANY = "Meridian Card Services (fictional)"
DISCLAIMER = (
    "Fictional company and data for portfolio purposes. Structure inspired by public "
    "information in American Express's Form 10-K and Oracle's public documentation. Not "
    "affiliated with or endorsed by American Express or Oracle. Contains no confidential data."
)
TENK = "https://www.sec.gov/Archives/edgar/data/4962/000000496226000080/axp-20251231.htm"
TENK_FIN = "https://www.sec.gov/Archives/edgar/data/4962/000000496226000080/"
ORA = "https://docs.oracle.com/en/cloud/saas/financials/26d/"

# ---- neutral styling (no third-party branding) ----
INK, SLATE, MUTED, LINE, PANEL = "1F2A30", "34495E", "6B7780", "D5DADD", "F3F5F6"
FONT = "Arial"
F = lambda **k: Font(name=FONT, size=k.pop("size", 10), color=k.pop("color", INK), **k)  # noqa: E731
F_BASE, F_BOLD = F(), F(bold=True)
F_HDR = F(bold=True, color="FFFFFF")
F_TITLE, F_SUB = F(size=18, bold=True), F(size=11, color=MUTED)
F_NOTE, F_SECTION = F(size=9, italic=True, color=MUTED), F(size=12, bold=True, color=SLATE)
F_INPUT = F(color="0000FF")
F_WARN = F(size=9, italic=True, bold=True, color="9C1C10")
FILL_HDR = PatternFill("solid", fgColor=SLATE)
FILL_ORA = PatternFill("solid", fgColor="E8ECEF")
FILL_PANEL = PatternFill("solid", fgColor=PANEL)
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color=LINE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
TOP_C = Alignment(horizontal="center", vertical="top")
AMT = '#,##0.00;(#,##0.00);"-"'
RAG = {"G": ("CDEBD3", "1E6B34"), "A": ("FCE7B2", "7A4F00"), "R": ("F8D0CC", "9C1C10"),
       "N": ("E4E7EA", "4A555D")}
HDR = 5  # header row on every data sheet; data starts on row 6


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def page(ws, title, subtitle, note=None):
    ws["A1"], ws["A2"] = title, subtitle
    ws["A1"].font, ws["A2"].font = F_TITLE, F_SUB
    ws.row_dimensions[1].height = 28
    if note:
        ws["A3"] = note
        ws["A3"].font = F_NOTE
    ws.sheet_view.showGridLines = False


def header(ws, row, cols, start=1, fill=FILL_HDR, font=F_HDR):
    for i, h in enumerate(cols):
        c = ws.cell(row=row, column=start + i, value=h)
        c.font, c.fill, c.alignment, c.border = font, fill, CENTER, BORDER
    ws.row_dimensions[row].height = 42


def body(ws, r1, r2, c1, c2, wrap=(), center=(), amount=()):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if cell.font == Font() or cell.font.name != FONT:
                cell.font = F_BASE
            cell.alignment = WRAP if c in wrap else (TOP_C if c in center else Alignment(vertical="top"))
            if c in amount:
                cell.number_format = AMT


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def rag_text(ws, ref, anchor, mapping):
    """mapping: {text: 'G'|'A'|'R'|'N'} -> fill whole range by the anchor cell's text."""
    for text, key in mapping.items():
        bg, fg = RAG[key]
        ws.conditional_formatting.add(ref, FormulaRule(
            formula=[f'{anchor}="{text}"'], stopIfTrue=True,
            fill=PatternFill("solid", fgColor=bg, bgColor=bg), font=Font(name=FONT, bold=True, color=fg)))


def rag_formula(ws, ref, rules):
    """rules: [(formula, 'G'|'A'|'R')] evaluated in order."""
    for formula, key in rules:
        bg, fg = RAG[key]
        ws.conditional_formatting.add(ref, FormulaRule(
            formula=[formula], stopIfTrue=True,
            fill=PatternFill("solid", fgColor=bg, bgColor=bg), font=Font(name=FONT, bold=True, color=fg)))


def fit_print(ws, tall=0):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, tall
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def text(v):
    return None if v is None or (isinstance(v, float) and pd.isna(v)) or v == "" else str(v)


def chart_style(ch, y_title):
    ch.roundedCorners = False
    ch.legend = None
    ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.title = y_title
    ch.y_axis.title.overlay = False
    ch.y_axis.majorGridlines = None
    ch.y_axis.numFmt = "0"


# ---------------------------------------------------------------------------
# sheets
# ---------------------------------------------------------------------------
def sheet_legacy(ws, coa, tb):
    page(ws, "Legacy COA", f"{COMPANY} · on-premises chart of accounts (3-part: Company-Account-CostCenter) "
         "with Oct-26 trial balance",
         "Data-quality issues are left in on purpose (duplicates, inactive accounts, inconsistent "
         "descriptions). Reporting Group = 10-K caption from the legacy FS-line crosswalk.")
    cols = ["Legacy Combination", "Company", "Account", "Cost Center", "Description", "Account Type",
            "Legacy Division", "Legacy FS Line", "Reporting Group (10-K caption)", "Status",
            "Last Activity", "TB Debit", "TB Credit", "Net Balance", "Data Quality Flag"]
    header(ws, HDR, cols)
    tbi = tb.set_index("Legacy Combination")
    seen, descs = set(), {}
    for d in coa.itertuples(index=False):
        descs.setdefault(d[0], []).append(d.Description)
    r = HDR
    for d in coa.itertuples(index=False):
        r += 1
        combo = d[0]
        first = combo not in seen
        seen.add(combo)
        vals = [combo, d.Company, d.Account, d[3], d.Description, d[5], d[6], d[7],
                LEGACY_FS_TO_GROUP[d[7]], d.Status, d[9]]
        for c, v in enumerate(vals, 1):
            ws.cell(row=r, column=c, value=v).number_format = "@"
        if first and combo in tbi.index:
            ws.cell(row=r, column=12, value=float(tbi.at[combo, "Debit"])).font = F_INPUT
            ws.cell(row=r, column=13, value=float(tbi.at[combo, "Credit"])).font = F_INPUT
        ws.cell(row=r, column=14, value=f"=L{r}-M{r}")
        flags = []
        if not first:
            flags.append("Duplicate row")
        elif len(descs[combo]) > 1:
            flags.append("Has duplicate")
        if d.Status == "Inactive":
            flags.append("Inactive")
        if d.Description != d.Description.strip() or (d.Description.isupper() and len(d.Description) > 6):
            flags.append("Inconsistent description")
        ws.cell(row=r, column=15, value="; ".join(flags) or None)
    last = r
    body(ws, HDR + 1, last, 1, 15, wrap=(5, 9, 15), center=(2, 3, 4, 10), amount=(12, 13, 14))
    for rr in range(HDR + 1, last + 1):
        for c in (12, 13):
            ws.cell(row=rr, column=c).font = F_INPUT
    t = last + 1
    ws.cell(row=t, column=1, value="Total")
    for c, L in ((12, "L"), (13, "M"), (14, "N")):
        ws.cell(row=t, column=c, value=f"=SUM({L}{HDR+1}:{L}{last})").number_format = AMT
    for c in range(1, 16):
        ws.cell(row=t, column=c).font = F_BOLD
        ws.cell(row=t, column=c).border = Border(top=Side(style="medium", color=SLATE))
    ws.cell(row=t + 1, column=1, value="Trial balance check")
    ws.cell(row=t + 1, column=1).font = F_BOLD
    ws.cell(row=t + 1, column=12, value=f'=IF(ABS(L{t}-M{t})<0.005,"BALANCED","OUT OF BALANCE")')
    rag_text(ws, f"L{t+1}", f"$L${t+1}", {"BALANCED": "G", "OUT OF BALANCE": "R"})
    rag_formula(ws, f"O{HDR+1}:O{last}", [
        (f'ISNUMBER(SEARCH("Duplicate",$O{HDR+1}))', "R"),
        (f'ISNUMBER(SEARCH("Inactive",$O{HDR+1}))', "N"),
        (f'LEN($O{HDR+1})>0', "A")])
    rag_text(ws, f"J{HDR+1}:J{last}", f"$J{HDR+1}", {"Inactive": "N"})
    widths(ws, [15, 9, 9, 9, 36, 11, 22, 11, 26, 9, 12, 17, 17, 17, 22])
    ws.freeze_panes = ws.cell(row=HDR + 1, column=2)
    ws.auto_filter.ref = f"A{HDR}:O{last}"
    fit_print(ws)
    return last


def sheet_target(ws, tcoa):
    page(ws, "Target COA", f"{COMPANY} · Oracle Cloud chart of accounts - 5 segments",
         "Segment labels follow Oracle's chart-of-accounts documentation. Line of Business values = "
         "reportable segments named in the FY2025 10-K (+ Corporate & Other, not a segment). "
         "Account Group = 10-K statement caption. Codes are fictional.")
    cols = ["Oracle Column", "Segment", "Segment Label", "Value", "Description", "Account Type",
            "Account Group (10-K caption)", "Parent Line of Business"]
    header(ws, HDR, cols)
    r = HDR
    for row in tcoa.itertuples(index=False):
        r += 1
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=text(v)).number_format = "@"
    last = r
    body(ws, HDR + 1, last, 1, 8, wrap=(5, 7, 8), center=(1, 4))
    # band each segment
    seg_fill = {"SEGMENT1": "FFFFFF", "SEGMENT2": "F3F5F6", "SEGMENT3": "FFFFFF",
                "SEGMENT4": "F3F5F6", "SEGMENT5": "FFFFFF"}
    for rr in range(HDR + 1, last + 1):
        f = PatternFill("solid", fgColor=seg_fill[ws.cell(row=rr, column=1).value])
        for c in range(1, 9):
            ws.cell(row=rr, column=c).fill = f
    ws.cell(row=last + 2, column=1, value=(
        "Cost Center has no segment label: Oracle's documentation reviewed did not confirm a "
        "'Cost Center' label name, so it is left blank rather than guessed.")).font = F_NOTE
    widths(ws, [12, 16, 25, 9, 40, 14, 30, 34])
    ws.freeze_panes = ws.cell(row=HDR + 1, column=1)
    ws.auto_filter.ref = f"A{HDR}:H{last}"
    fit_print(ws)
    return last


def sheet_mapping(ws, mp, legacy_last):
    page(ws, "Mapping", f"{COMPANY} · legacy combination → Oracle Cloud combination",
         "Edit blue segment cells to remap; Target Combination, Legacy Net Balance and Flag recalculate. "
         "Flags: OK · RETIRED · UNMAPPED · MISSING SEGMENT · RETIRED WITH BALANCE.")
    cols = ["Legacy Combination", "Legacy Description", "Legacy Status", "Legacy Net Balance",
            "Target Combination", "Segment1 Company", "Segment2 Line of Business", "Segment3 Cost Center",
            "Segment4 Natural Account", "Segment5 Intercompany", "Mapping Type", "Notes", "Flag"]
    header(ws, HDR, cols)
    L = f"'Legacy COA'!${{}}${HDR+1}:${{}}${legacy_last}"
    r = HDR
    for row in mp.itertuples(index=False):
        r += 1
        ws.cell(row=r, column=1, value=row[0]).number_format = "@"
        ws.cell(row=r, column=2, value=row[1])
        ws.cell(row=r, column=3, value=row[2])
        ws.cell(row=r, column=4, value=f"=SUMIFS({L.format('N','N')},{L.format('A','A')},A{r})")
        ws.cell(row=r, column=5, value=f'=IF(COUNTBLANK(F{r}:J{r})>0,"",F{r}&"-"&G{r}&"-"&H{r}&"-"&I{r}&"-"&J{r})')
        for i, s in enumerate(["Segment1", "Segment2", "Segment3", "Segment4", "Segment5"]):
            c = ws.cell(row=r, column=6 + i, value=text(getattr(row, s)))
            c.number_format = "@"
        ws.cell(row=r, column=11, value=row[8]).number_format = "@"
        ws.cell(row=r, column=12, value=row[9])
        ws.cell(row=r, column=13, value=(
            f'=IF(K{r}="Retired",IF(ABS(D{r})>0.004,"RETIRED WITH BALANCE","RETIRED"),'
            f'IF(K{r}="Unmapped","UNMAPPED",IF(COUNTBLANK(F{r}:J{r})>0,"MISSING SEGMENT","OK")))'))
    last = r
    body(ws, HDR + 1, last, 1, 13, wrap=(2, 12), center=(3, 6, 7, 8, 9, 10, 11, 13), amount=(4,))
    for rr in range(HDR + 1, last + 1):
        for c in range(6, 11):
            ws.cell(row=rr, column=c).font = F_INPUT
    rag_text(ws, f"M{HDR+1}:M{last}", f"$M{HDR+1}", {
        "OK": "G", "RETIRED": "N", "UNMAPPED": "R", "MISSING SEGMENT": "R", "RETIRED WITH BALANCE": "R"})
    rag_text(ws, f"A{HDR+1}:A{last}", f"$M{HDR+1}", {
        "UNMAPPED": "R", "MISSING SEGMENT": "R", "RETIRED WITH BALANCE": "R"})
    dv = DataValidation(type="list", formula1='"1:1,Many:1,Retired,Unmapped"')
    ws.add_data_validation(dv)
    dv.add(f"K{HDR+1}:K{last}")
    t = last + 2
    ws.cell(row=t, column=1, value="Disposition summary").font = F_SECTION
    for i, mt in enumerate(["1:1", "Many:1", "Retired", "Unmapped"]):
        ws.cell(row=t + 1 + i, column=1, value=mt).number_format = "@"
        ws.cell(row=t + 1 + i, column=2,
                value=f'=SUMPRODUCT(--($K${HDR+1}:$K${last}=A{t+1+i}))')
    ws.cell(row=t + 5, column=1, value="Total").font = F_BOLD
    ws.cell(row=t + 5, column=2, value=f"=SUM(B{t+1}:B{t+4})").font = F_BOLD
    body(ws, t + 1, t + 5, 1, 2, center=(2,))
    widths(ws, [15, 34, 10, 17, 25, 10, 12, 11, 11, 14, 11, 52, 20])
    ws.freeze_panes = ws.cell(row=HDR + 1, column=2)
    ws.auto_filter.ref = f"A{HDR}:M{last}"
    fit_print(ws)
    return last


FBDI_ORACLE = ["STATUS", "LEDGER_NAME", "ACCOUNTING_DATE", "USER_JE_SOURCE_NAME",
               "USER_JE_CATEGORY_NAME", "CURRENCY_CODE", "ACTUAL_FLAG", "SEGMENT1", "SEGMENT2",
               "SEGMENT3", "SEGMENT4", "SEGMENT5", "ENTERED_DR", "ENTERED_CR", "REFERENCE10"]


def sheet_fbdi(ws, j, target_last):
    page(ws, "FBDI Journal Template", "SIMPLIFIED VERSION - NOT THE OFFICIAL ORACLE TEMPLATE")
    ws["A2"].font = F(size=11, bold=True, color="9C1C10")
    ws["A3"] = ("Columns follow the order of Oracle's GL_INTERFACE control file (GlInterface.ctl); "
                "only SEGMENT1-5 are used and optional columns are omitted. STATUS = NEW and ACTUAL_FLAG = A "
                "per Oracle's GL_INTERFACE reference. Grey columns Q-R are helpers for reconciliation "
                "and are not part of the upload.")
    ws["A3"].font = F_NOTE
    header(ws, 4, FBDI_ORACLE, fill=FILL_ORA, font=F(size=8, bold=True, color=SLATE))
    ws.row_dimensions[4].height = 26
    cols = list(j.columns)
    header(ws, HDR, cols)
    header(ws, HDR, ["Net (helper)", "Account Group (helper)"], start=17, fill=FILL_ORA,
           font=F(bold=True, color=SLATE))
    tgt_d = f"'Target COA'!$D${HDR+1}:$D${target_last}"
    tgt_g = f"'Target COA'!$G${HDR+1}:$G${target_last}"
    r = HDR
    for row in j.itertuples(index=False):
        r += 1
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=float(v) if c in (13, 14) else text(v))
            if c not in (13, 14):
                cell.number_format = "@"
        ws.cell(row=r, column=17, value=f"=M{r}-N{r}")
        ws.cell(row=r, column=18, value=f'=IFERROR(INDEX({tgt_g},MATCH(K{r},{tgt_d},0)),"NOT FOUND")')
    last = r
    body(ws, HDR + 1, last, 1, 15, wrap=(15,), center=(1, 3, 6, 7, 8, 9, 10, 11, 12), amount=(13, 14))
    body(ws, HDR + 1, last, 17, 18, amount=(17,))
    for rr in range(HDR + 1, last + 1):
        for c in (17, 18):
            ws.cell(row=rr, column=c).fill = PatternFill("solid", fgColor="F7F8F9")
            ws.cell(row=rr, column=c).font = F(color=MUTED)
    t = last + 1
    ws.cell(row=t, column=12, value="Total")
    ws.cell(row=t, column=13, value=f"=SUM(M{HDR+1}:M{last})")
    ws.cell(row=t, column=14, value=f"=SUM(N{HDR+1}:N{last})")
    ws.cell(row=t + 1, column=12, value="Dr - Cr")
    ws.cell(row=t + 1, column=13, value=f"=M{t}-N{t}")
    for rr in (t, t + 1):
        for c in (12, 13, 14):
            ws.cell(row=rr, column=c).font = F_BOLD
            ws.cell(row=rr, column=c).number_format = AMT
    rag_formula(ws, f"M{t+1}", [(f"ABS($M${t+1})<0.005", "G"), (f"ABS($M${t+1})>=0.005", "R")])
    rag_text(ws, f"R{HDR+1}:R{last}", f"$R{HDR+1}", {"NOT FOUND": "R"})
    widths(ws, [8, 22, 11, 16, 17, 9, 7, 9, 9, 9, 10, 9, 17, 17, 52, 2, 19, 28])
    ws.freeze_panes = ws.cell(row=HDR + 1, column=1)
    ws.auto_filter.ref = f"A{HDR}:O{last}"
    fit_print(ws)
    return last


def sheet_validation(ws, errs, map_last, fb_last, leg_last):
    page(ws, "Validation Errors", f"{COMPANY} · pre-load control checks and exception log",
         "Control checks are live formulas. An exception BLOCKS the load when the affected combination "
         "carries a balance; zero-balance findings are logged as errors to clear before go-live.")
    ws["A4"] = "Control checks"
    ws["A4"].font = F_SECTION
    header(ws, HDR, ["#", "Control check", "", "", "Result"])
    ws.merge_cells(start_row=HDR, start_column=2, end_row=HDR, end_column=4)
    FB = "'FBDI Journal Template'!"
    fr = lambda c: f"{FB}${c}${HDR+1}:${c}${fb_last}"  # noqa: E731
    MP = lambda c: f"Mapping!${c}${HDR+1}:${c}${map_last}"  # noqa: E731
    LG = lambda c: f"'Legacy COA'!${c}${HDR+1}:${c}${leg_last}"  # noqa: E731
    checks = [
        ("Legacy trial balance is balanced (debits = credits)",
         f'=IF(ABS(SUM({LG("L")})-SUM({LG("M")}))<0.005,"PASS","FAIL")'),
        ("Journal is balanced - all companies",
         f'=IF(ABS(SUM({fr("M")})-SUM({fr("N")}))<0.005,"PASS","FAIL")'),
        ("Journal is balanced - Company 101 (primary balancing segment)",
         f'=IF(ABS(SUMIFS({fr("M")},{fr("H")},"101")-SUMIFS({fr("N")},{fr("H")},"101"))<0.005,"PASS","FAIL")'),
        ("Journal is balanced - Company 201 (primary balancing segment)",
         f'=IF(ABS(SUMIFS({fr("M")},{fr("H")},"201")-SUMIFS({fr("N")},{fr("H")},"201"))<0.005,"PASS","FAIL")'),
        ("Every journal line has all 5 segments",
         f'=IF(COUNTBLANK({FB}$H${HDR+1}:$L${fb_last})=0,"PASS","FAIL")'),
        ("Every natural account exists in the Target COA",
         f'=IF(COUNTIF({fr("R")},"NOT FOUND")=0,"PASS","FAIL")'),
        ("No balance on unmapped, retired or incomplete combinations",
         f'=IF(SUMPRODUCT(({MP("M")}<>"OK")*(ABS({MP("D")})>0.004))=0,"PASS","FAIL")'),
        ("STATUS = NEW and ACTUAL_FLAG = A on every line",
         f'=IF(AND(COUNTIF({fr("A")},"NEW")=COUNTA({fr("A")}),COUNTIF({fr("G")},"A")=COUNTA({fr("A")})),"PASS","FAIL")'),
    ]
    for i, (label, f) in enumerate(checks, 1):
        r = HDR + i
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=label)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=5, value=f)
    c_last = HDR + len(checks)
    body(ws, HDR + 1, c_last, 1, 5, wrap=(2,), center=(1, 5))
    for rr in range(HDR + 1, c_last + 1):
        ws.row_dimensions[rr].height = 28
    rag_text(ws, f"E{HDR+1}:E{c_last}", f"$E{HDR+1}", {"PASS": "G", "FAIL": "R"})

    s = c_last + 2
    ws.cell(row=s, column=1, value="Exception log").font = F_SECTION
    h = s + 1
    cols = ["Issue ID", "Check", "Severity", "Blocks Load", "Legacy Combination", "Target Combination",
            "Combination Balance", "Detail", "Resolution"]
    header(ws, h, cols)
    r = h
    for e in errs.itertuples(index=False):
        r += 1
        ws.cell(row=r, column=1, value=e[0])
        ws.cell(row=r, column=2, value=e[1])
        ws.cell(row=r, column=3, value=e[2])
        ws.cell(row=r, column=4, value=f'=IF(AND(C{r}="Error",ABS(G{r})>0.004),"Yes","No")')
        ws.cell(row=r, column=5, value=text(e[4])).number_format = "@"
        ws.cell(row=r, column=6, value=text(e[5])).number_format = "@"
        ws.cell(row=r, column=7, value=(f'=IFERROR(INDEX({MP("D")},MATCH(E{r},{MP("A")},0)),0)'
                                        if text(e[4]) else float(e[6])))
        ws.cell(row=r, column=8, value=e[7])
        ws.cell(row=r, column=9, value=e[8])
    e_last = r
    body(ws, h + 1, e_last, 1, 9, wrap=(8, 9), center=(1, 3, 4), amount=(7,))
    for rr in range(h + 1, e_last + 1):
        ws.row_dimensions[rr].height = 45
    rag_text(ws, f"C{h+1}:C{e_last}", f"$C{h+1}", {"Error": "R", "Warning": "A"})
    rag_text(ws, f"D{h+1}:D{e_last}", f"$D{h+1}", {"Yes": "R", "No": "G"})
    widths(ws, [8, 18, 10, 11, 17, 22, 15, 60, 44])
    ws.freeze_panes = "A6"
    fit_print(ws, tall=1)
    return (HDR + 1, c_last), (h + 1, e_last)


def sheet_recon(ws, leg_last, fb_last, map_last):
    page(ws, "Reconciliation", f"{COMPANY} · legacy trial balance vs Oracle Cloud journal (Oct-26)",
         "Legacy side uses the OLD system's own division and FS-line hierarchies; Cloud side uses "
         "Segment2 and the target account group. Variance = Cloud - Legacy. Tolerance: 0.00.")
    ws["A4"], ws["E4"] = "Overall result:", None
    ws["A4"].font = F_SECTION
    cols = ["Level", "Item", "Legacy key", "Cloud key", "Legacy Amount", "Cloud Amount", "Variance", "Result"]
    header(ws, HDR, cols)
    LG = lambda c: f"'Legacy COA'!${c}${HDR+1}:${c}${leg_last}"  # noqa: E731
    FB = lambda c: f"'FBDI Journal Template'!${c}${HDR+1}:${c}${fb_last}"  # noqa: E731
    rows = [
        ("Overall", "Total debits", "TB Debit", "Entered Dr", f"=SUM({LG('L')})", f"=SUM({FB('M')})"),
        ("Overall", "Total credits", "TB Credit", "Entered Cr", f"=SUM({LG('M')})", f"=SUM({FB('N')})"),
        ("Overall", "Net (debits - credits)", "Net", "Net", f"=SUM({LG('N')})", f"=SUM({FB('Q')})"),
    ]
    div = {"10": "CONSUMER", "20": "COMMERCIAL", "30": "INTERNATIONAL", "40": "MERCHANT & NETWORK",
           "90": "CORPORATE"}
    for code, name in LOB.items():
        rows.append(("Line of Business", name, div[code], code, None, None))
    for g in dict.fromkeys(v[2] for v in NATURAL_ACCOUNT.values()):
        rows.append(("Account group", g, g, g, None, None))
    r = HDR
    for lvl, item, lk, ck, lf, cf in rows:
        r += 1
        ws.cell(row=r, column=1, value=lvl)
        ws.cell(row=r, column=2, value=item)
        ws.cell(row=r, column=3, value=lk).number_format = "@"
        ws.cell(row=r, column=4, value=ck).number_format = "@"
        if lvl == "Line of Business":
            lf = f"=SUMIFS({LG('N')},{LG('G')},C{r})"
            cf = f"=SUMIFS({FB('Q')},{FB('I')},D{r})"
        elif lvl == "Account group":
            lf = f"=SUMIFS({LG('N')},{LG('I')},C{r})"
            cf = f"=SUMIFS({FB('Q')},{FB('R')},D{r})"
        ws.cell(row=r, column=5, value=lf)
        ws.cell(row=r, column=6, value=cf)
        ws.cell(row=r, column=7, value=f"=ROUND(F{r}-E{r},2)")
        ws.cell(row=r, column=8, value=f'=IF(ABS(G{r})<0.005,"PASS","FAIL")')
    last = r
    body(ws, HDR + 1, last, 1, 8, center=(1, 8), amount=(5, 6, 7))
    for rr in range(HDR + 1, last + 1):  # band levels
        if ws.cell(row=rr, column=1).value == "Line of Business":
            for c in range(1, 9):
                ws.cell(row=rr, column=c).fill = FILL_PANEL
    ws["C4"] = f'=IF(COUNTIF(H{HDR+1}:H{last},"FAIL")=0,"PASS","FAIL")'
    ws["C4"].font = F(size=14, bold=True)
    ws["C4"].alignment = CENTER
    ws["D4"] = f'=COUNTIF(H{HDR+1}:H{last},"PASS")&" of "&COUNTA(H{HDR+1}:H{last})&" checks pass"'
    ws["D4"].font = F_NOTE
    rag_text(ws, "C4", "$C$4", {"PASS": "G", "FAIL": "R"})
    rag_text(ws, f"H{HDR+1}:H{last}", f"$H{HDR+1}", {"PASS": "G", "FAIL": "R"})
    rag_formula(ws, f"G{HDR+1}:G{last}", [(f"ABS($G{HDR+1})>=0.005", "R")])

    s = last + 2
    ws.cell(row=s, column=1, value="Exceptions").font = F_SECTION
    ex = [
        ("Reconciliation lines with a variance", f'=COUNTIF(H{HDR+1}:H{last},"FAIL")'),
        ("Legacy trial-balance lines", f"=COUNT({LG('L')})"),
        ("Cloud journal lines (after Many:1 consolidation)", f"=COUNTA({FB('A')})"),
        ("Legacy combinations not carried forward (Retired / Unmapped / Missing segment)",
         f"=COUNTA(Mapping!$A${HDR+1}:$A${map_last})-COUNTIF(Mapping!$M${HDR+1}:$M${map_last},\"OK\")"),
        ("…of which carry a balance (would break the reconciliation)",
         f"=SUMPRODUCT((Mapping!$M${HDR+1}:$M${map_last}<>\"OK\")*(Mapping!$A${HDR+1}:$A${map_last}<>\"\")"
         f"*(ABS(Mapping!$D${HDR+1}:$D${map_last})>0.004))"),
    ]
    for i, (lbl, f) in enumerate(ex, 1):
        ws.cell(row=s + i, column=1, value=lbl)
        ws.merge_cells(start_row=s + i, start_column=1, end_row=s + i, end_column=4)
        ws.cell(row=s + i, column=5, value=f)
    body(ws, s + 1, s + len(ex), 1, 5, center=(5,))
    rag_formula(ws, f"E{s+1}", [(f"$E${s+1}=0", "G"), (f"$E${s+1}>0", "R")])
    rag_formula(ws, f"E{s+5}", [(f"$E${s+5}=0", "G"), (f"$E${s+5}>0", "R")])
    widths(ws, [17, 36, 27, 27, 19, 19, 14, 9])
    ws.freeze_panes = ws.cell(row=HDR + 1, column=1)
    fit_print(ws, tall=1)
    return "C4"


FITGAP = [
    ("Record international card portfolios in local currency and report consolidated results in USD.",
     "Multi-currency", "Fit",
     "Standard: load Daily Rates (spreadsheet or web service), run the Revaluation Process on foreign-"
     "currency balances, use Reporting Currencies / Translation for the USD view.",
     "Standard feature", "Y", "Rates come from one automated source; revaluation runs are logged "
     "instead of manual FX spreadsheets.", "High"),
    ("Intercompany transactions between legal entities and lines of business balance automatically.",
     "Intercompany", "Fit",
     "Configure Intercompany Balancing Rules on the primary and second balancing segments; use the "
     "Intercompany Segment (101 / 201) to identify the counterparty.",
     "Configuration", "Y", "Unbalanced intercompany journals cannot post; removes manual true-up "
     "entries and month-end intercompany breaks.", "High"),
    ("Allocate Corporate & Other costs (technology, occupancy) to lines of business by driver.",
     "Allocations", "Fit",
     "Standard allocation rules generating allocation journals each period; cross-ledger "
     "allocations if a second ledger is added.", "Standard feature", "Y",
     "Rule-based, repeatable allocations replace spreadsheet calculations.", "Medium"),
    ("Journals above amount thresholds, or from manual sources, need approval; preparers cannot "
     "approve their own entries.", "Journal approval", "Fit",
     "Journal Approval rules built with Oracle Business Process Management (amount and source "
     "conditions, approval hierarchy).", "Configuration", "Y",
     "Enforced approval with recorded approval history on every journal (preventive SOX-style control).",
     "High"),
    ("Control period open / close by ledger and see close status in one place.", "Period close", "Fit",
     "Accounting Calendar with period open and close per ledger; Close Monitor for close status; "
     "Intercompany Period Status for intercompany cut-off.", "Standard feature", "Y",
     "Closed periods block back-dated postings; close status visible to leadership.", "High"),
    ("Full audit trail from source system to GL balance.", "Audit trail", "Partial",
     "Journal source, category and reference columns (e.g. REFERENCE10 line description) on every "
     "imported journal plus approval history. Audit reporting for setup changes to be confirmed in design.",
     "Configuration + process", "Y",
     "Every converted or imported line traces back to its legacy account or source file.", "High"),
    ("Prevent invalid account combinations (e.g. revenue on a balance-sheet-only cost center).",
     "Controls", "Fit", "Cross-Validation Rules to block invalid combinations at entry and import.",
     "Configuration", "Y", "Preventive control replaces detective month-end review of mis-coded entries.",
     "High"),
    ("Management reporting by reportable segment (P&L and key balances).", "Reporting", "Fit",
     "Line of Business as the second balancing segment, Account Hierarchies, and Financial Reporting / "
     "Smart View on the GL balances cubes.", "Standard feature", "Y",
     "One source for segment results; no off-system re-mapping in spreadsheets.", "High"),
    ("Restrict users to the segment data they own (e.g. International team sees only ICS).",
     "Security", "Fit", "Segment Value Security rules by Line of Business and Company.",
     "Configuration", "Y", "Least-privilege data access; supports segregation of duties.", "Medium"),
    ("Month-end accruals reverse automatically and approved journals post on a schedule.",
     "Journal processing", "Fit", "Journal Reversals (AutoReverse) and AutoPost criteria sets.",
     "Standard feature", "Y", "No missed or duplicated reversals; posting follows a defined schedule.",
     "Medium"),
    ("Load high-volume daily feeds from the card-processing platform.", "Integration", "Partial",
     "Summarise at source and load with Journal Import (FBDI file, then the Import Journals process) on "
     "a schedule; assess subledger accounting for transaction-level detail.",
     "Process change + integration", "Y",
     "Automated load with control totals replaces manual uploads.", "High"),
    ("Calculate the Card Member rewards liability and credit loss reserve (CECL-style) estimates.",
     "Accounting estimates", "Gap",
     "Keep actuarial and credit-loss models outside the GL; post results via Journal Import routed "
     "through journal approval. No GL customisation.", "Process change", "Y",
     "Model output is reviewed and approved before posting; journal references the model run.", "High"),
    ("Reconcile merchant settlement and network clearing accounts.", "Reconciliation", "Partial",
     "Clearing Accounts Reconciliation for GL clearing accounts; balance-sheet certification process "
     "for other accounts.", "Standard feature + process", "Y",
     "Unreconciled items are aged and visible; fewer manual reconciliations.", "Medium"),
    ("Keep 7+ years of legacy history available without slowing the new ledger.", "Data retention",
     "Partial", "Convert opening balances only (as in this simulation); keep detail in an archive or data "
     "warehouse; use Accounting Data Archive and Purge for ongoing retention.", "Process change", "Y",
     "Smaller conversion scope lowers risk; documented retention policy.", "Low"),
    ("Reproduce a legacy regulator-specific report layout exactly.", "Reporting", "Gap",
     "Would need a custom report built as an extension. Recommendation: agree a standard layout from "
     "Financial Reporting instead.", "Extension", "N",
     "If built, report logic is documented and version-controlled.", "Low"),
]


def sheet_fitgap(ws):
    page(ws, "Fit-Gap", f"{COMPANY} · GL and close requirements vs standard Oracle Cloud functionality",
         "Approach: standard functionality first, then configuration, then process change; extensions "
         "only where unavoidable. Oracle feature names as used in Oracle's public GL documentation.")
    cols = ["ID", "Requirement", "Process Area", "Fit / Partial / Gap", "Proposed Solution",
            "Solution Type", "Avoids Customization? (Y/N)", "Control Improvement", "Priority"]
    header(ws, HDR, cols)
    r = HDR
    for i, row in enumerate(FITGAP, 1):
        r += 1
        ws.cell(row=r, column=1, value=f"FG{i:02d}")
        for c, v in enumerate(row, 2):
            ws.cell(row=r, column=c, value=v)
        ws.row_dimensions[r].height = 62
    last = r
    body(ws, HDR + 1, last, 1, 9, wrap=(2, 3, 5, 6, 8), center=(1, 4, 7, 9))
    rag_text(ws, f"D{HDR+1}:D{last}", f"$D{HDR+1}", {"Fit": "G", "Partial": "A", "Gap": "R"})
    rag_text(ws, f"G{HDR+1}:G{last}", f"$G{HDR+1}", {"Y": "G", "N": "R"})
    rag_text(ws, f"I{HDR+1}:I{last}", f"$I{HDR+1}", {"High": "R", "Medium": "A", "Low": "N"})
    for col, vals in (("D", "Fit,Partial,Gap"), ("G", "Y,N"), ("I", "High,Medium,Low")):
        dv = DataValidation(type="list", formula1=f'"{vals}"')
        ws.add_data_validation(dv)
        dv.add(f"{col}{HDR+1}:{col}{last}")
    widths(ws, [7, 40, 15, 11, 52, 15, 13, 40, 9])
    ws.freeze_panes = ws.cell(row=HDR + 1, column=3)
    ws.auto_filter.ref = f"A{HDR}:I{last}"
    fit_print(ws, tall=1)
    return last


def sheet_summary(ws, map_last, val_rng, fb_last, fg_last):
    page(ws, "Oracle GL Migration Simulator – Summary",
         f"{COMPANY} · Card issuer GL migration, Oracle on-premises to Oracle Cloud · Oct-26 balance conversion")
    ws.merge_cells("A3:J3")
    ws["A3"] = DISCLAIMER
    ws["A3"].font = F_WARN
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 30
    (c1, c2), (e1, e2) = val_rng
    MP = lambda c: f"Mapping!${c}${HDR+1}:${c}${map_last}"  # noqa: E731
    VE = lambda c: f"'Validation Errors'!${c}${e1}:${c}${e2}"  # noqa: E731
    FG = lambda c: f"'Fit-Gap'!${c}${HDR+1}:${c}${fg_last}"  # noqa: E731
    ws["A5"] = "Migration KPIs"
    ws["A5"].font = F_SECTION
    tiles = [
        ("A", "B", "% mapped", f'=IFERROR(COUNTIF({MP("M")},"OK")/(COUNTA({MP("A")})-COUNTIF({MP("K")},"Retired")),0)', "0.0%"),
        ("C", "D", "Validation errors", f'=COUNTIF({VE("C")},"Error")', "0"),
        ("E", "F", "Blocking errors", f'=COUNTIF({VE("D")},"Yes")', "0"),
        ("G", "H", "Reconciliation", "=Reconciliation!$C$4", "@"),
        ("I", "J", "# Gaps", f'=COUNTIF({FG("D")},"Gap")', "0"),
    ]
    for a, b, label, f, fmt in tiles:
        ws.merge_cells(f"{a}6:{b}6")
        ws.merge_cells(f"{a}7:{b}7")
        ws[f"{a}6"], ws[f"{a}7"] = label, f
        ws[f"{a}6"].font = F(size=9, bold=True, color=MUTED)
        ws[f"{a}7"].font = F(size=22, bold=True)
        ws[f"{a}6"].alignment = ws[f"{a}7"].alignment = CENTER
        ws[f"{a}7"].number_format = fmt
        for col in (a, b):
            for rr in (6, 7):
                ws[f"{col}{rr}"].fill = FILL_PANEL
                ws[f"{col}{rr}"].border = BORDER
    ws.row_dimensions[7].height = 42
    rag_formula(ws, "A7:B7", [("$A$7>=0.98", "G"), ("$A$7>=0.9", "A"), ("$A$7<0.9", "R")])
    rag_formula(ws, "C7:D7", [("$C$7=0", "G"), ("$C$7>0", "A")])
    rag_formula(ws, "E7:F7", [("$E$7=0", "G"), ("$E$7>0", "R")])
    rag_text(ws, "G7:H7", "$G$7", {"PASS": "G", "FAIL": "R"})
    rag_formula(ws, "I7:J7", [("$I$7=0", "G"), ("$I$7<=2", "A"), ("$I$7>2", "R")])
    ws["A8"] = ("% mapped = OK combinations ÷ (all legacy combinations - Retired). Validation errors are "
                "zero-balance findings to clear before go-live; blocking errors would stop the load.")
    ws["A8"].font = F_NOTE

    ws["A10"] = "Load readiness"
    ws["A10"].font = F_SECTION
    detail = [
        ("Control checks passed", f'=COUNTIF(\'Validation Errors\'!$E${c1}:$E${c2},"PASS")&" of "&COUNTA(\'Validation Errors\'!$E${c1}:$E${c2})'),
        ("Legacy combinations (unique)", f"=COUNTA({MP('A')})"),
        ("Journal lines in FBDI file", f"=COUNTA('FBDI Journal Template'!$A${HDR+1}:$A${fb_last})"),
        ("Total debits converted", f"=SUM('FBDI Journal Template'!$M${HDR+1}:$M${fb_last})"),
        ("Warnings (non-blocking)", f'=COUNTIF({VE("C")},"Warning")'),
        ("Partial fits", f'=COUNTIF({FG("D")},"Partial")'),
        ("Requirements avoiding customization",
         f'=COUNTIF({FG("G")},"Y")&" of "&COUNTA({FG("G")})'),
        ("Ready to import?", f'=IF(AND($E$7=0,$G$7="PASS",COUNTIF(\'Validation Errors\'!$E${c1}:$E${c2},"FAIL")=0),"READY","NOT READY")'),
    ]
    for i, (lbl, f) in enumerate(detail):
        r = 11 + i
        ws.cell(row=r, column=1, value=lbl)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=f)
        c.number_format = AMT if "debits" in lbl else "General"
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    body(ws, 11, 18, 1, 5, center=(4,))
    for c in range(1, 6):
        ws.cell(row=18, column=c).font = F_BOLD
    rag_text(ws, "D18:E18", "$D$18", {"READY": "G", "NOT READY": "R"})

    # chart data
    ws["A21"] = "Mapping disposition"
    ws["A21"].font = F_SECTION
    header(ws, 22, ["Mapping Type", "Combinations"])
    for i, mt in enumerate(["1:1", "Many:1", "Retired", "Unmapped"]):
        r = 23 + i
        ws.cell(row=r, column=1, value=mt).number_format = "@"
        ws.cell(row=r, column=2, value=f"=SUMPRODUCT(--({MP('K')}=A{r}))")
    body(ws, 23, 26, 1, 2, center=(1, 2))
    ws["D21"] = "Fit-gap result"
    ws["D21"].font = F_SECTION
    header(ws, 22, ["Result", "Requirements"], start=4)
    for i, v in enumerate(["Fit", "Partial", "Gap"]):
        r = 23 + i
        ws.cell(row=r, column=4, value=v)
        ws.cell(row=r, column=5, value=f'=COUNTIF({FG("D")},D{r})')
    body(ws, 23, 25, 4, 5, center=(4, 5))

    ch = BarChart()
    ch.type, ch.title = "col", "Legacy combinations by mapping type"
    ch.add_data(Reference(ws, min_col=2, min_row=22, max_row=26), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=23, max_row=26))
    ch.series[0].graphicalProperties.solidFill = "5B8DB8"
    ch.series[0].graphicalProperties.line.solidFill = "5B8DB8"
    chart_style(ch, "Combinations")
    ch.height, ch.width = 7.5, 13
    ws.add_chart(ch, "G10")

    ch2 = BarChart()
    ch2.type, ch2.title = "col", "Fit-gap: 15 GL & close requirements"
    ch2.add_data(Reference(ws, min_col=5, min_row=22, max_row=25), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=4, min_row=23, max_row=25))
    from openpyxl.chart.series import DataPoint
    for idx, colour in enumerate(["3C9A5F", "E0A030", "C8553D"]):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = colour
        pt.graphicalProperties.line.solidFill = colour
        ch2.series[0].dPt.append(pt)
    chart_style(ch2, "Requirements")
    ch2.height, ch2.width = 7.5, 13
    ws.add_chart(ch2, "G26")
    widths(ws, [14, 14, 14, 14, 14, 4, 14, 14, 14, 14])
    fit_print(ws, tall=1)


def sheet_about(ws):
    page(ws, "About & Sources", "What is real, what is fictional, and where the real parts come from")
    ws.merge_cells("A4:D4")
    ws["A4"] = DISCLAIMER
    ws["A4"].font = F(size=11, bold=True, color="9C1C10")
    ws["A4"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[4].height = 46
    r = 6

    def section(title):
        nonlocal r
        ws.cell(row=r, column=1, value=title).font = F_SECTION
        r += 1

    def line(k, v, link=None):
        nonlocal r
        ws.cell(row=r, column=1, value=k).font = F_BOLD
        ws.cell(row=r, column=1).alignment = WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(row=r, column=2, value=v)
        c.alignment = WRAP
        c.font = F_BASE
        if link:
            c.hyperlink = link
            c.font = F(color="1F5F8B", underline="single")
        ws.row_dimensions[r].height = max(15, 15 * (len(str(v)) // 95 + 1))
        r += 1

    section("Source 1 - American Express Company, Form 10-K, fiscal year ended December 31, 2025")
    line("Filed", "February 6, 2026 (SEC accession no. 0000004962-26-000080)")
    line("Document", TENK, TENK)
    line("Financial statements", "Consolidated Statements of Income and Consolidated Balance Sheets, read "
         "from the filing's financial report pages (R3, R6) in the same EDGAR folder.", TENK_FIN)
    line("Used", "Reportable segment names (MD&A, Business Introduction); revenue, interest, provision and "
         "expense captions (Consolidated Statements of Income); asset, liability and equity captions "
         "(Consolidated Balance Sheets). Names only - no amounts were taken.")
    r += 1
    section("Source 2 - Oracle public documentation (Oracle Fusion Cloud Financials, release 26D)")
    line("FBDI Journal Import", ORA + "oefbf/journalimport-3096.html", ORA + "oefbf/journalimport-3096.html")
    line("GL_INTERFACE table", ORA + "oedmf/glinterface-10390.html", ORA + "oedmf/glinterface-10390.html")
    line("Control file", "https://www.oracle.com/webfolder/technetwork/docs/fbdi-26d/fbdi/controlfiles/GlInterface.ctl",
         "https://www.oracle.com/webfolder/technetwork/docs/fbdi-26d/fbdi/controlfiles/GlInterface.ctl")
    line("Chart of accounts", ORA + "faigl/create-a-chart-of-accounts-structure-and-instance.html",
         ORA + "faigl/create-a-chart-of-accounts-structure-and-instance.html")
    line("Balancing segments", ORA + "faigl/overview-of-balancing-segments.html",
         ORA + "faigl/overview-of-balancing-segments.html")
    line("GL features", "Implementing Enterprise Structures and General Ledger guide (" + ORA + "faigl/), "
         "page titles for Intercompany Balancing Rules, Journal Approval, Cross-Validation Rules, Segment "
         "Value Security, Account Hierarchies, Revaluation Process, Reporting Currencies, Daily Rates, "
         "Close Monitor, Journal Reversals, AutoPost, Clearing Accounts Reconciliation, Accounting Data "
         "Archive and Purge, Financial Reporting and Smart View.")
    line("Checked", "November 2026.")
    r += 1
    section("Real (public structure only)")
    for k, v in [
        ("Line of Business values", "U.S. Consumer Services; Commercial Services; International Card Services; "
         "Global Merchant and Network Services (the four reportable operating segments). Corporate & Other is "
         "also used - the 10-K says it is NOT a reportable segment."),
        ("Account groups", "10-K captions: Discount revenue; Net card fees; Service fees and other revenue; "
         "Total interest income; Total interest expense; Provisions for credit losses; Card Member rewards; "
         "Business development; Card Member services; Marketing; Salaries and employee benefits; Other, net; "
         "Cash and cash equivalents; Card Member receivables; Card Member loans; Premises and equipment; "
         "Other assets; Customer deposits; Accounts payable; Other liabilities; Long-term debt; Common shares; "
         "Retained earnings."),
        ("Oracle terms", "GL_INTERFACE column order and names; STATUS = NEW; ACTUAL_FLAG = A; REFERENCE10 = "
         "journal entry line description; Import Journals process; JournalImportTemplate.xlsm; segment labels "
         "Primary Balancing Segment, Second Balancing Segment, Natural Account Segment, Intercompany Segment."),
    ]:
        line(k, v)
    r += 1
    section("Fictional (everything else)")
    for k, v in [
        ("Company & ledger", "Meridian Card Services and the 'MCS USD Primary Ledger' are invented."),
        ("Accounts & codes", "All legacy and target account codes, cost centers, descriptions and mappings."),
        ("Balances", "All trial balance amounts. No figures from the 10-K are used."),
        ("Requirements", "The 15 fit-gap requirements, solutions, priorities and control notes."),
        ("Placement choice", "Holding the Membership Rewards liability under 'Other liabilities' is a "
         "fictional design choice, not taken from the 10-K."),
    ]:
        line(k, v)
    r += 1
    section("Deliberately left out (could not be verified)")
    for k, v in [
        ("Processed revenue", "Not used: the FY2025 10-K states it was consolidated into Service fees and "
         "other revenue from Q1 2025."),
        ("Cost Center label", "No segment label is shown for Cost Center: the Oracle pages reviewed did not "
         "confirm a label name."),
        ("FBDI template", "The FBDI sheet is a SIMPLIFIED VERSION, NOT THE OFFICIAL ORACLE TEMPLATE. Optional "
         "columns (e.g. DATE_CREATED, SEGMENT6-30, converted amounts, other REFERENCE columns) are omitted."),
    ]:
        line(k, v)
    widths(ws, [24, 40, 30, 50])


def main() -> None:
    coa = pd.read_csv(LEGACY / "legacy_coa.csv", dtype=str)
    tb = pd.read_csv(LEGACY / "legacy_trial_balance.csv", dtype={"Legacy Combination": str})
    tcoa = pd.read_csv(CLOUD / "target_coa.csv", dtype=str, keep_default_na=False)
    mp = pd.read_csv(CLOUD / "coa_mapping.csv", dtype=str, keep_default_na=False)
    mp = mp[["Legacy Combination", "Legacy Description", "Legacy Status", "Segment1", "Segment2",
             "Segment3", "Segment4", "Segment5", "Mapping Type", "Notes"]]
    j = pd.read_csv(CLOUD / "fbdi_gl_interface.csv", dtype=str, keep_default_na=False)
    errs = pd.read_csv(CLOUD / "validation_errors.csv", dtype=str, keep_default_na=False)
    recon = pd.read_csv(CLOUD / "reconciliation.csv")
    assert (recon["Result"] == "PASS").all(), "run reconcile.py first; reconciliation must pass"

    wb = Workbook()
    names = ["Summary", "Legacy COA", "Target COA", "Mapping", "Fit-Gap", "FBDI Journal Template",
             "Validation Errors", "Reconciliation", "About & Sources"]
    wb.active.title = names[0]
    for n in names[1:]:
        wb.create_sheet(n)
    leg_last = sheet_legacy(wb["Legacy COA"], coa, tb)
    tgt_last = sheet_target(wb["Target COA"], tcoa)
    map_last = sheet_mapping(wb["Mapping"], mp, leg_last)
    fg_last = sheet_fitgap(wb["Fit-Gap"])
    fb_last = sheet_fbdi(wb["FBDI Journal Template"], j, tgt_last)
    val_rng = sheet_validation(wb["Validation Errors"], errs, map_last, fb_last, leg_last)
    sheet_recon(wb["Reconciliation"], leg_last, fb_last, map_last)
    sheet_summary(wb["Summary"], map_last, val_rng, fb_last, fg_last)
    sheet_about(wb["About & Sources"])
    tabs = ["34495E", "8A969E", "2A7F78", "2A7F78", "C8553D", "5B8DB8", "9C1C10", "3C9A5F", "6B7780"]
    for n, col in zip(names, tabs):
        wb[n].sheet_properties.tabColor = col
    wb.properties.title = "Oracle GL Migration Simulator – Card Issuer, On-Prem to Cloud"
    wb.properties.creator = "Portfolio project"
    wb.properties.description = DISCLAIMER
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
