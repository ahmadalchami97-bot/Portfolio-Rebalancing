#!/usr/bin/env python3
"""
Generator for: Family_Office_Portfolio_Allocator.xlsx

A KWD-base, institutional strategic asset-allocation & rebalancing workbook.
Native Excel only (formulas, named ranges, tables, charts) - no VBA / macros.

Run:  python3 build_workbook.py
"""

import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import column_index_from_string

OUTFILE = "Family_Office_Portfolio_Allocator.xlsx"

# ---------------------------------------------------------------- sheet names
S1 = "1. Settings"
S2 = "2. Portfolio Inputs"
S3 = "3. Dashboard"
S4 = "4. Capital Deployment"
S5 = "5. Rebalance"
S6 = "6. Performance"

# ---------------------------------------------------------------- palette
NAVY        = "1F3864"   # primary accent (titles)
GRAY_HEADER = "D9D9D9"   # light gray section / table headers
GRAY_LIGHT  = "F2F2F2"   # very light banding
GRAY_TOTAL  = "E7E6E6"   # total rows
YELLOW      = "FFF6CC"   # input cells
WHITE       = "FFFFFF"
DARK        = "1F2937"   # near-black text
GREY_TXT    = "595959"
RED         = "C00000"
GREEN       = "375623"
BORDER      = "BFBFBF"

# ---------------------------------------------------------------- fonts
F_TITLE   = Font(name="Calibri", size=16, bold=True, color=WHITE)
F_SUB     = Font(name="Calibri", size=9,  italic=True, color="DCE3F0")
F_SECTION = Font(name="Calibri", size=11, bold=True, color=NAVY)
F_HEAD    = Font(name="Calibri", size=10, bold=True, color=DARK)
F_LABEL   = Font(name="Calibri", size=10, color=DARK)
F_LABELB  = Font(name="Calibri", size=10, bold=True, color=DARK)
F_DATA    = Font(name="Calibri", size=10, color=DARK)
F_INPUT   = Font(name="Calibri", size=10, color="7F6000")
F_NOTE    = Font(name="Calibri", size=9,  italic=True, color=GREY_TXT)
F_BIG     = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_KPI     = Font(name="Calibri", size=11, bold=True, color=DARK)

# ---------------------------------------------------------------- fills
FILL_TITLE   = PatternFill("solid", fgColor=NAVY)
FILL_SECTION = PatternFill("solid", fgColor=GRAY_HEADER)
FILL_HEAD    = PatternFill("solid", fgColor=GRAY_HEADER)
FILL_INPUT   = PatternFill("solid", fgColor=YELLOW)
FILL_BAND    = PatternFill("solid", fgColor=GRAY_LIGHT)
FILL_TOTAL   = PatternFill("solid", fgColor=GRAY_TOTAL)
FILL_BOX     = PatternFill("solid", fgColor=GRAY_LIGHT)
FILL_WHITE   = PatternFill("solid", fgColor=WHITE)

# ---------------------------------------------------------------- borders
_thin = Side(style="thin", color=BORDER)
BOX   = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
NOBORDER = Border()

# ---------------------------------------------------------------- number formats
KWD   = '#,##0'
KWDc  = '#,##0;[Red]-#,##0'
FC    = '#,##0'
FX    = '0.000000'
PCT   = '0.0%'
PCTc  = '0.0%;[Red]-0.0%'
PCT2  = '0.00%'
DATE  = 'yyyy-mm-dd'

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center")
LEFTW  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT  = Alignment(horizontal="right",  vertical="center")

ASSETS = ["Gold", "Swiss Funds", "S&P 500 ETFs", "US Government Bonds"]


# ============================================================ helpers
def C(ws, coord, value=None, font=F_DATA, fill=None, align=None, fmt=None,
      border=BOX, locked=True):
    cell = ws[coord]
    if value is not None:
        cell.value = value
    cell.font = font
    if fill is not None:
        cell.fill = fill
    if align is not None:
        cell.alignment = align
    else:
        cell.alignment = Alignment(vertical="center")
    if fmt is not None:
        cell.number_format = fmt
    if border is not None:
        cell.border = border
    cell.protection = Protection(locked=locked)
    return cell


def inp(ws, coord, value=None, fmt=None, align=RIGHT):
    """Yellow input cell (visually unlocked)."""
    return C(ws, coord, value=value, font=F_INPUT, fill=FILL_INPUT,
             align=align, fmt=fmt, border=BOX, locked=False)


def merge(ws, rng, value, font, fill=None, align=LEFT, border=None):
    ws.merge_cells(rng)
    first = rng.split(":")[0]
    c = ws[first]
    c.value = value
    c.font = font
    if fill is not None:
        c.fill = fill
    c.alignment = align
    if border is not None:
        for row in ws[rng]:
            for cc in row:
                cc.border = border
    if fill is not None:                       # paint full merged band
        for row in ws[rng]:
            for cc in row:
                cc.fill = fill
    return c


def title_bar(ws, rng_title, title, rng_sub, sub):
    merge(ws, rng_title, title, F_TITLE, FILL_TITLE, LEFT)
    ws.row_dimensions[int(rng_title.split(":")[0][1:])].height = 26
    merge(ws, rng_sub, sub, F_SUB, FILL_TITLE, LEFT)


def section(ws, rng, text):
    merge(ws, rng, text, F_SECTION, FILL_SECTION, LEFT)


def note(ws, rng, text):
    merge(ws, rng, text, F_NOTE, None, LEFTW)


def headers(ws, row, start_col, names):
    for i, name in enumerate(names):
        C(ws, f"{chr(64+start_col+i)}{row}", name, font=F_HEAD,
          fill=FILL_HEAD, align=CENTER, border=BOX)


def kpi_box(ws, label_coord, label, value_range, formula, fmt=KWD):
    C(ws, label_coord, label, font=F_LABELB, fill=FILL_SECTION, align=LEFT)
    ws.merge_cells(value_range)
    first = value_range.split(":")[0]
    c = ws[first]
    c.value = formula
    c.font = F_BIG
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.number_format = fmt
    for row in ws[value_range]:
        for cc in row:
            cc.fill = FILL_BOX
            cc.border = BOX
    return c


def addname(wb, name, sheet, ref):
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!{ref}"))


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def add_table(ws, name, ref, style="TableStyleLight1"):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(
        name=style, showRowStripes=True, showColumnStripes=False,
        showFirstColumn=False, showLastColumn=False)
    ws.add_table(t)


def finish(ws, freeze="A4"):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    ws.freeze_panes = freeze
    ws.column_dimensions["A"].width = 2.5


# ============================================================ build
wb = Workbook()
wb.calculation.fullCalcOnLoad = True   # force Excel/LibreOffice to recalc on open

ws1 = wb.active
ws1.title = S1
ws2 = wb.create_sheet(S2)
ws3 = wb.create_sheet(S3)
ws4 = wb.create_sheet(S4)
ws5 = wb.create_sheet(S5)
ws6 = wb.create_sheet(S6)


# ------------------------------------------------------------ SHEET 1 : SETTINGS
def build_settings(ws):
    title_bar(ws, "B1:F1", "Settings & Assumptions",
              "B2:F2", "Central assumptions, FX rates, target allocation and policy. "
                       "Yellow cells are inputs; all other cells are calculated.")

    # Base currency
    section(ws, "B4:D4", "Base Currency")
    C(ws, "B5", "Base reporting currency", font=F_LABEL, align=LEFT)
    inp(ws, "C5", "KWD", align=CENTER)

    # Exchange rates
    section(ws, "B7:D7", "Exchange Rates")
    note(ws, "B8:D8", "Enter the KWD value of 1 unit of each foreign currency. "
                      "Add other currencies in the spare rows as needed.")
    headers(ws, 9, 2, ["Currency", "Rate (KWD per 1 unit)"])
    fx = [("USD", 0.307000), ("CHF", 0.340000), ("KWD", 1.000000), ("", None), ("", None)]
    for i, (cur, rate) in enumerate(fx):
        r = 10 + i
        if cur:
            inp(ws, f"B{r}", cur, align=CENTER)
        else:
            inp(ws, f"B{r}", None, align=CENTER)
        inp(ws, f"C{r}", rate, fmt=FX)

    # Strategic target allocation
    section(ws, "B16:D16", "Strategic Target Allocation")
    headers(ws, 17, 2, ["Asset Class", "Target %"])
    targets = [0.20, 0.20, 0.40, 0.20]
    for i, a in enumerate(ASSETS):
        r = 18 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        inp(ws, f"C{r}", targets[i], fmt=PCT)
    C(ws, "B22", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C22", "=SUM(C18:C21)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    merge(ws, "B23:D23",
          '=IF(ROUND(Target_Total,4)=1,"✓ Target allocation balanced (100%)",'
          '"⚠ Target allocation must equal 100% (currently "&TEXT(Target_Total,"0.0%")&")")',
          F_LABELB, None, LEFT)

    # Rebalancing parameters
    section(ws, "B25:D25", "Rebalancing Parameters")
    C(ws, "B26", "Tolerance Band (%)", font=F_LABEL, align=LEFT)
    inp(ws, "C26", 0.03, fmt=PCT)
    note(ws, "B27:D27", "No rebalance is recommended unless an asset class deviates "
                        "from target by more than this band.")

    # Minimum action threshold
    section(ws, "B29:D29", "Minimum Action Threshold")
    C(ws, "B30", "Minimum Action (KWD)", font=F_LABEL, align=LEFT)
    inp(ws, "C30", 100000, fmt=KWD)
    note(ws, "B31:D31", "Allocation differences smaller than this amount are ignored.")

    # Expected returns
    section(ws, "B33:D33", "Expected Return Assumptions")
    note(ws, "B34:D34", "Used only for projections and scenario analysis (Performance sheet).")
    headers(ws, 35, 2, ["Asset Class", "Expected Annual Return"])
    exp = [0.05, 0.04, 0.08, 0.03]
    for i, a in enumerate(ASSETS):
        r = 36 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        inp(ws, f"C{r}", exp[i], fmt=PCT)

    # Instructions
    section(ws, "B41:F41", "Instructions")
    steps = [
        "1.  Update market values and currencies on the 'Portfolio Inputs' sheet.",
        "2.  Update FX rates and policy assumptions on this sheet.",
        "3.  Review current allocation and gaps on the 'Dashboard'.",
        "4.  Use 'Capital Deployment' to allocate new capital to underweight classes.",
        "5.  Use 'Rebalance' for strategic capital-movement recommendations.",
        "6.  Use 'Performance' for returns, CAGR, scenarios and growth projections.",
    ]
    for i, s in enumerate(steps):
        merge(ws, f"B{42+i}:F{42+i}", s, F_LABEL, None, LEFTW)

    # conditional warning colour on the 100% status
    ws.conditional_formatting.add(
        "B23",
        FormulaRule(formula=["ROUND(Target_Total,4)<>1"],
                    font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(
        "B23",
        FormulaRule(formula=["ROUND(Target_Total,4)=1"],
                    font=Font(color=GREEN, bold=True)))

    widths(ws, {"B": 30, "C": 20, "D": 18, "E": 14, "F": 26})
    finish(ws, "A3")


# ------------------------------------------------------------ SHEET 2 : INPUTS
def build_inputs(ws):
    title_bar(ws, "B1:F1", "Portfolio Inputs",
              "B2:F2", "Single source of portfolio market values. Update foreign-currency "
                       "values and currency; KWD is computed automatically.")
    kpi_box(ws, "B4", "Portfolio Total Value (KWD)", "C4:E4", "=Portfolio_Total", KWD)

    headers(ws, 6, 2, ["Asset Class", "Market Value (FC)", "Currency",
                       "FX Rate (KWD/FC)", "Market Value (KWD)"])
    data = [("Gold", 3000000, "USD"), ("Swiss Funds", 4000000, "CHF"),
            ("S&P 500 ETFs", 8000000, "USD"), ("US Government Bonds", 5000000, "USD")]
    for i, (a, mv, cur) in enumerate(data):
        r = 7 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        inp(ws, f"C{r}", mv, fmt=FC)
        inp(ws, f"D{r}", cur, align=CENTER)
        C(ws, f"E{r}", f"=IFERROR(VLOOKUP(D{r},FX_Table,2,FALSE),0)", fmt=FX, align=RIGHT)
        C(ws, f"F{r}", f"=C{r}*E{r}", fmt=KWD, align=RIGHT)
    # total row (outside table)
    C(ws, "B11", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    for col in ("C", "D", "E"):
        C(ws, f"{col}11", None, fill=FILL_TOTAL)
    C(ws, "F11", "=SUM(F7:F10)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)

    add_table(ws, "tblInputs", "B6:F10")

    dv = DataValidation(type="list", formula1=f"='{S1}'!$B$10:$B$14", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("D7:D10")

    note(ws, "B13:F14", "FX Rate is looked up automatically from the Exchange Rates table "
                        "on the Settings sheet, based on the currency selected. "
                        "Market Value (KWD) = Market Value (FC) × FX Rate.")

    widths(ws, {"B": 24, "C": 20, "D": 14, "E": 18, "F": 22})
    finish(ws, "A7")


# ------------------------------------------------------------ SHEET 3 : DASHBOARD
def build_dashboard(ws):
    title_bar(ws, "B1:G1", "Allocation Dashboard",
              "B2:G2", "Current portfolio value, allocation versus target, and allocation gaps.")
    kpi_box(ws, "B3", "Portfolio Total Value (KWD)", "C3:E3", "=Portfolio_Total", KWD)

    headers(ws, 5, 2, ["Asset Class", "Current Value (KWD)", "Current %",
                       "Target %", "Difference %"])
    names = ["=MV_Gold", "=MV_Swiss", "=MV_SP500", "=MV_Bonds"]
    tgt = ["=Target_Gold", "=Target_Swiss", "=Target_SP500", "=Target_Bonds"]
    for i, a in enumerate(ASSETS):
        r = 6 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", names[i], fmt=KWD, align=RIGHT)
        C(ws, f"D{r}", f"=IFERROR(C{r}/Portfolio_Total,0)", fmt=PCT, align=RIGHT)
        C(ws, f"E{r}", tgt[i], fmt=PCT, align=RIGHT)
        C(ws, f"F{r}", f"=D{r}-E{r}", fmt=PCTc, align=RIGHT)
    C(ws, "B10", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C10", "=SUM(C6:C9)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "D10", "=SUM(D6:D9)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    C(ws, "E10", "=SUM(E6:E9)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    C(ws, "F10", "=D10-E10", font=F_LABELB, fill=FILL_TOTAL, fmt=PCTc, align=RIGHT)
    add_table(ws, "tblAlloc", "B5:F9")

    # Underweight
    section(ws, "B12:D12", "Underweight Assets (below target)")
    headers(ws, 13, 2, ["Asset Class", "Allocation Gap %", "Allocation Gap (KWD)"])
    for i in range(4):
        ar = 6 + i      # alloc-table row
        r = 14 + i
        C(ws, f"B{r}", f'=IF(D{ar}<E{ar},B{ar},"")', font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", f'=IF(D{ar}<E{ar},E{ar}-D{ar},"")', fmt=PCT, align=RIGHT)
        C(ws, f"D{r}", f'=IF(D{ar}<E{ar},(E{ar}-D{ar})*Portfolio_Total,"")', fmt=KWD, align=RIGHT)
    C(ws, "B18", "Total underweight", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C18", None, fill=FILL_TOTAL)
    C(ws, "D18", "=SUM(D14:D17)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)

    # Overweight
    section(ws, "B20:D20", "Overweight Assets (above target)")
    headers(ws, 21, 2, ["Asset Class", "Allocation Gap %", "Allocation Gap (KWD)"])
    for i in range(4):
        ar = 6 + i
        r = 22 + i
        C(ws, f"B{r}", f'=IF(D{ar}>E{ar},B{ar},"")', font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", f'=IF(D{ar}>E{ar},D{ar}-E{ar},"")', fmt=PCT, align=RIGHT)
        C(ws, f"D{r}", f'=IF(D{ar}>E{ar},(D{ar}-E{ar})*Portfolio_Total,"")', fmt=KWD, align=RIGHT)
    C(ws, "B26", "Total overweight", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C26", None, fill=FILL_TOTAL)
    C(ws, "D26", "=SUM(D22:D25)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)

    # difference colour
    ws.conditional_formatting.add("F6:F9",
        CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED)))
    ws.conditional_formatting.add("F6:F9",
        CellIsRule(operator="greaterThan", formula=["0"], font=Font(color=GREEN)))

    # Charts ----------------------------------------------------------------
    pie1 = PieChart()
    pie1.title = "Current Allocation"
    pie1.height = 7.5
    pie1.width = 11
    pie1.add_data(Reference(ws, min_col=3, min_row=5, max_row=9), titles_from_data=True)
    pie1.set_categories(Reference(ws, min_col=2, min_row=6, max_row=9))
    pie1.dataLabels = DataLabelList()
    pie1.dataLabels.showPercent = True
    ws.add_chart(pie1, "H5")

    pie2 = PieChart()
    pie2.title = "Target Allocation"
    pie2.height = 7.5
    pie2.width = 11
    pie2.add_data(Reference(ws, min_col=5, min_row=5, max_row=9), titles_from_data=True)
    pie2.set_categories(Reference(ws, min_col=2, min_row=6, max_row=9))
    pie2.dataLabels = DataLabelList()
    pie2.dataLabels.showPercent = True
    ws.add_chart(pie2, "H21")

    widths(ws, {"B": 24, "C": 20, "D": 14, "E": 12, "F": 14, "G": 3})
    finish(ws, "A5")


# ------------------------------------------------------------ SHEET 4 : DEPLOYMENT
def build_deployment(ws):
    title_bar(ws, "B1:J1", "Capital Deployment",
              "B2:J2", "Determine where new capital should be deployed. Underweight asset "
                       "classes are funded toward target first, then by target weights.")
    C(ws, "B3", "New Capital Available (KWD)", font=F_LABELB, fill=FILL_SECTION, align=LEFT)
    inp(ws, "C3", 500000, fmt=KWD)
    C(ws, "E3", "Current Portfolio (KWD)", font=F_LABELB, fill=FILL_SECTION, align=LEFT)
    C(ws, "F3", "=Portfolio_Total", fmt=KWD, align=RIGHT, fill=FILL_BOX)
    C(ws, "H3", "Total After Deployment (KWD)", font=F_LABELB, fill=FILL_SECTION, align=LEFT)
    C(ws, "I3", "=Portfolio_Total+New_Capital", fmt=KWD, align=RIGHT, fill=FILL_BOX)

    headers(ws, 5, 2, ["Asset Class", "Current Value (KWD)", "Target %",
                       "Target Value @ New Total", "Gap to Target (KWD)",
                       "Suggested Allocation (KWD)", "Value After (KWD)",
                       "Allocation After %", "Current %"])
    mv = ["=MV_Gold", "=MV_Swiss", "=MV_SP500", "=MV_Bonds"]
    tgt = ["=Target_Gold", "=Target_Swiss", "=Target_SP500", "=Target_Bonds"]
    for i, a in enumerate(ASSETS):
        r = 6 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", mv[i], fmt=KWD, align=RIGHT)
        C(ws, f"D{r}", tgt[i], fmt=PCT, align=RIGHT)
        C(ws, f"E{r}", f"=D{r}*New_Total", fmt=KWD, align=RIGHT)
        C(ws, f"F{r}", f"=MAX(0,E{r}-C{r})", fmt=KWD, align=RIGHT)
        C(ws, f"G{r}",
          f"=IF(New_Capital<=0,0,IF($F$10<=0,New_Capital*D{r},"
          f"IF(New_Capital<=$F$10,New_Capital*F{r}/$F$10,"
          f"F{r}+(New_Capital-$F$10)*D{r})))",
          fmt=KWD, align=RIGHT)
        C(ws, f"H{r}", f"=C{r}+G{r}", fmt=KWD, align=RIGHT)
        C(ws, f"I{r}", f"=IFERROR(H{r}/New_Total,0)", fmt=PCT, align=RIGHT)
        C(ws, f"J{r}", f"=IFERROR(C{r}/Portfolio_Total,0)", fmt=PCT, align=RIGHT)
    # total
    C(ws, "B10", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    for col, f in [("C", "=SUM(C6:C9)"), ("E", "=SUM(E6:E9)"), ("F", "=SUM(F6:F9)"),
                   ("G", "=SUM(G6:G9)"), ("H", "=SUM(H6:H9)")]:
        C(ws, f"{col}10", f, font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "D10", None, fill=FILL_TOTAL)
    C(ws, "I10", "=SUM(I6:I9)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    C(ws, "J10", "=SUM(J6:J9)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    add_table(ws, "tblDeploy", "B5:J9")

    # Recommended deployment list
    section(ws, "B12:D12", "Recommended Capital Deployment")
    rows = [("New Capital Available", "=New_Capital"),
            ("Gold", "=G6"), ("Swiss Funds", "=G7"),
            ("S&P 500 ETFs", "=G8"), ("US Government Bonds", "=G9")]
    for i, (lab, f) in enumerate(rows):
        r = 13 + i
        fnt = F_LABELB if i == 0 else F_LABEL
        C(ws, f"B{r}", lab, font=fnt, align=LEFT)
        C(ws, f"C{r}", None)
        C(ws, f"D{r}", f, font=fnt, fmt=KWD, align=RIGHT)
    C(ws, "B18", "Total Allocated", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C18", None, fill=FILL_TOTAL)
    C(ws, "D18", "=SUM(D14:D17)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)

    # Expected allocation after deployment
    section(ws, "B20:F20", "Expected Allocation After Deployment")
    headers(ws, 21, 2, ["Asset Class", "Value After (KWD)", "Allocation After %", "Target %",
                        "Remaining Gap to Target (KWD)"])
    for i, a in enumerate(ASSETS):
        r = 22 + i
        ar = 6 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", f"=H{ar}", fmt=KWD, align=RIGHT)
        C(ws, f"D{r}", f"=I{ar}", fmt=PCT, align=RIGHT)
        C(ws, f"E{r}", f"=D{ar}", fmt=PCT, align=RIGHT)
        C(ws, f"F{r}", f"=MAX(0,E{r}*New_Total-C{r})", fmt=KWD, align=RIGHT)
    C(ws, "B26", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C26", "=SUM(C22:C25)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "D26", "=SUM(D22:D25)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    C(ws, "E26", "=SUM(E22:E25)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    C(ws, "F26", "=SUM(F22:F25)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)

    # before/after chart
    bar = BarChart()
    bar.type = "col"
    bar.title = "Allocation: Current % vs After Deployment %"
    bar.height = 7.5
    bar.width = 13
    bar.add_data(Reference(ws, min_col=10, min_row=5, max_row=9), titles_from_data=True)  # J Current %
    bar.add_data(Reference(ws, min_col=9, min_row=5, max_row=9), titles_from_data=True)   # I After %
    bar.set_categories(Reference(ws, min_col=2, min_row=6, max_row=9))
    bar.y_axis.numFmt = "0%"
    bar.y_axis.delete = False
    bar.x_axis.delete = False
    ws.add_chart(bar, "B28")

    widths(ws, {"B": 22, "C": 18, "D": 12, "E": 18, "F": 16, "G": 18,
                "H": 16, "I": 14, "J": 12})
    finish(ws, "A5")


# ------------------------------------------------------------ SHEET 5 : REBALANCE
def build_rebalance(ws):
    title_bar(ws, "B1:H1", "Rebalance Analysis",
              "B2:H2", "Strategic rebalancing - capital movements between asset classes. "
                       "No trade tickets, no units, no security-level recommendations.")
    kpi_box(ws, "B3", "Portfolio Total Value (KWD)", "C3:E3", "=Portfolio_Total", KWD)

    headers(ws, 5, 2, ["Asset Class", "Current Value (KWD)", "Target %",
                       "Target Value (KWD)", "Difference (KWD)", "Deviation %", "Action"])
    mv = ["=MV_Gold", "=MV_Swiss", "=MV_SP500", "=MV_Bonds"]
    tgt = ["=Target_Gold", "=Target_Swiss", "=Target_SP500", "=Target_Bonds"]
    for i, a in enumerate(ASSETS):
        r = 6 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", mv[i], fmt=KWD, align=RIGHT)
        C(ws, f"D{r}", tgt[i], fmt=PCT, align=RIGHT)
        C(ws, f"E{r}", f"=D{r}*Portfolio_Total", fmt=KWD, align=RIGHT)
        C(ws, f"F{r}", f"=E{r}-C{r}", fmt=KWDc, align=RIGHT)
        C(ws, f"G{r}", f"=IFERROR(C{r}/Portfolio_Total,0)-D{r}", fmt=PCTc, align=RIGHT)
        C(ws, f"H{r}",
          f'=IF(ABS(G{r})<=Tol_Band,"Hold",IF(F{r}>0,"Increase","Reduce"))',
          font=F_LABELB, align=CENTER)
    C(ws, "B10", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C10", "=SUM(C6:C9)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "D10", None, fill=FILL_TOTAL)
    C(ws, "E10", "=SUM(E6:E9)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "F10", "=SUM(F6:F9)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWDc, align=RIGHT)
    C(ws, "G10", None, fill=FILL_TOTAL)
    C(ws, "H10", None, fill=FILL_TOTAL)
    add_table(ws, "tblRebal", "B5:H9")

    # Recommendation table
    section(ws, "B12:C12", "Recommended Capital Movements")
    headers(ws, 13, 2, ["Recommendation", "Amount (KWD)"])
    for i in range(4):
        rr = 6 + i
        r = 14 + i
        C(ws, f"B{r}",
          f'=IF(H{rr}="Hold","Hold "&B{rr}&" (within tolerance)",H{rr}&" "&B{rr})',
          font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", f'=IF(H{rr}="Hold","",ABS(F{rr}))', fmt=KWD, align=RIGHT)
    C(ws, "B18", "Total Capital To Reallocate", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C18", '=SUMPRODUCT((H6:H9="Increase")*(F6:F9))',
      font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)

    # Narrative
    section(ws, "B20:E20", "Suggested Capital Movements")
    note(ws, "B21:E21", "Act on the following to move the portfolio toward target allocation:")
    for i in range(4):
        rr = 6 + i
        r = 22 + i
        merge(ws, f"B{r}:E{r}",
              f'=IF(H{rr}="Hold","•  "&B{rr}&" — hold (within tolerance band)",'
              f'"•  "&H{rr}&" "&B{rr}&" by "&TEXT(ABS(F{rr}),"#,##0")&" KWD")',
              F_LABEL, None, LEFT)

    # After rebalancing
    section(ws, "B27:E27", "Portfolio Allocation After Rebalancing")
    headers(ws, 28, 2, ["Asset Class", "Value After (KWD)", "Allocation After %", "Target %"])
    for i, a in enumerate(ASSETS):
        rr = 6 + i
        r = 29 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        C(ws, f"C{r}", f'=IF(COUNTIF($H$6:$H$9,"Hold")=4,C{rr},E{rr})', fmt=KWD, align=RIGHT)
        C(ws, f"D{r}", f"=IFERROR(C{r}/$C$33,0)", fmt=PCT, align=RIGHT)
        C(ws, f"E{r}", f"=D{rr}", fmt=PCT, align=RIGHT)
    C(ws, "B33", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C33", "=SUM(C29:C32)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "D33", "=SUM(D29:D32)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    C(ws, "E33", "=SUM(E29:E32)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    note(ws, "B34:E35", "A rebalance moves capital from overweight to underweight classes; the "
                        "portfolio total is unchanged and no new capital is added. When any class "
                        "breaches the tolerance band, the portfolio is returned to its strategic "
                        "target weights (shown above); the 'Action' column flags which classes have drifted.")

    # action colour
    ws.conditional_formatting.add("H6:H9",
        FormulaRule(formula=['$H6="Increase"'], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add("H6:H9",
        FormulaRule(formula=['$H6="Reduce"'], font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add("H6:H9",
        FormulaRule(formula=['$H6="Hold"'], font=Font(color=GREY_TXT)))

    # current vs target chart
    bar = BarChart()
    bar.type = "col"
    bar.title = "Current Value vs Target Value (KWD)"
    bar.height = 8
    bar.width = 13
    bar.add_data(Reference(ws, min_col=3, min_row=5, max_row=9), titles_from_data=True)  # current
    bar.add_data(Reference(ws, min_col=5, min_row=5, max_row=9), titles_from_data=True)  # target
    bar.set_categories(Reference(ws, min_col=2, min_row=6, max_row=9))
    bar.y_axis.numFmt = "#,##0"
    bar.y_axis.delete = False
    bar.x_axis.delete = False
    ws.add_chart(bar, "J5")

    widths(ws, {"B": 24, "C": 18, "D": 12, "E": 18, "F": 16, "G": 12, "H": 12, "I": 3})
    finish(ws, "A5")


# ------------------------------------------------------------ SHEET 6 : PERFORMANCE
def build_performance(ws):
    title_bar(ws, "B1:H1", "Performance & Simulation",
              "B2:H2", "Performance measurement, historical allocation scenarios and "
                       "growth projections.")
    C(ws, "B3", "Valuation Date", font=F_LABELB, fill=FILL_SECTION, align=LEFT)
    inp(ws, "C3", datetime.date(2026, 6, 16), fmt=DATE, align=CENTER)

    # SECTION A
    section(ws, "B5:H5", "A.  Current Performance (per Asset Class)")
    headers(ws, 6, 2, ["Asset Class", "Initial Value (KWD)", "Initial Date",
                       "Current Value (KWD)", "Gain / Loss (KWD)", "Return %", "CAGR"])
    initials = [(800000, datetime.date(2021, 1, 1)),
                (1200000, datetime.date(2021, 1, 1)),
                (1800000, datetime.date(2021, 1, 1)),
                (1500000, datetime.date(2021, 1, 1))]
    mv = ["=MV_Gold", "=MV_Swiss", "=MV_SP500", "=MV_Bonds"]
    for i, a in enumerate(ASSETS):
        r = 7 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        inp(ws, f"C{r}", initials[i][0], fmt=KWD)
        inp(ws, f"D{r}", initials[i][1], fmt=DATE, align=CENTER)
        C(ws, f"E{r}", mv[i], fmt=KWD, align=RIGHT)
        C(ws, f"F{r}", f"=E{r}-C{r}", fmt=KWDc, align=RIGHT)
        C(ws, f"G{r}", f"=IFERROR(E{r}/C{r}-1,0)", fmt=PCTc, align=RIGHT)
        C(ws, f"H{r}", f'=IFERROR((E{r}/C{r})^(1/YEARFRAC(D{r},Val_Date,1))-1,"")',
          fmt=PCTc, align=RIGHT)
    C(ws, "B11", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C11", "=SUM(C7:C10)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "D11", None, fill=FILL_TOTAL)
    C(ws, "E11", "=SUM(E7:E10)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "F11", "=E11-C11", font=F_LABELB, fill=FILL_TOTAL, fmt=KWDc, align=RIGHT)
    C(ws, "G11", "=IFERROR(E11/C11-1,0)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCTc, align=RIGHT)
    C(ws, "H11", '=IFERROR((E11/C11)^(1/YEARFRAC(MIN(D7:D10),Val_Date,1))-1,"")',
      font=F_LABELB, fill=FILL_TOTAL, fmt=PCTc, align=RIGHT)
    add_table(ws, "tblPerfA", "B6:H10")

    # SECTION B
    section(ws, "B13:D13", "B.  Portfolio Performance")
    bsec = [("Portfolio Initial Value (KWD)", "=SUM(C7:C10)", KWD),
            ("Portfolio Current Value (KWD)", "=Portfolio_Total", KWD),
            ("Portfolio Gain / Loss (KWD)", "=C15-C14", KWDc),
            ("Portfolio Return %", "=IFERROR(C15/C14-1,0)", PCTc),
            ("Portfolio Inception Date", "=MIN(D7:D10)", DATE),
            ("Portfolio CAGR", '=IFERROR((C15/C14)^(1/YEARFRAC(C18,Val_Date,1))-1,"")', PCTc)]
    for i, (lab, f, fmt) in enumerate(bsec):
        r = 14 + i
        C(ws, f"B{r}", lab, font=F_LABELB, align=LEFT)
        C(ws, f"C{r}", f, font=F_KPI, fmt=fmt, align=RIGHT)

    # SECTION C
    section(ws, "B21:H21", "C.  Historical Allocation Scenario")
    csec = [("Start Date", datetime.date(2016, 1, 1), DATE),
            ("End Date", datetime.date(2026, 1, 1), DATE),
            ("Starting Capital (KWD)", 1000000, KWD)]
    for i, (lab, val, fmt) in enumerate(csec):
        r = 22 + i
        C(ws, f"B{r}", lab, font=F_LABEL, align=LEFT)
        inp(ws, f"C{r}", val, fmt=fmt, align=(CENTER if fmt == DATE else RIGHT))
    C(ws, "B25", "Period (years)", font=F_LABEL, align=LEFT)
    C(ws, "C25", "=YEARFRAC(C22,C23,1)", fmt="0.00", align=RIGHT)

    headers(ws, 27, 2, ["Asset Class", "Allocation %", "Assumed Annual Return %",
                        "Growth Factor", "Start Allocation (KWD)", "Ending Value (KWD)",
                        "Contribution to Return %"])
    alloc = [0.20, 0.20, 0.40, 0.20]
    aret = [0.06, 0.04, 0.09, 0.03]
    for i, a in enumerate(ASSETS):
        r = 28 + i
        C(ws, f"B{r}", a, font=F_LABEL, align=LEFT)
        inp(ws, f"C{r}", alloc[i], fmt=PCT)
        inp(ws, f"D{r}", aret[i], fmt=PCT)
        C(ws, f"E{r}", f"=(1+D{r})^$C$25", fmt="0.000", align=RIGHT)
        C(ws, f"F{r}", f"=C{r}*$C$24", fmt=KWD, align=RIGHT)
        C(ws, f"G{r}", f"=F{r}*E{r}", fmt=KWD, align=RIGHT)
        C(ws, f"H{r}", f"=IFERROR((G{r}-F{r})/$C$24,0)", fmt=PCTc, align=RIGHT)
    C(ws, "B32", "Total", font=F_LABELB, fill=FILL_TOTAL, align=LEFT)
    C(ws, "C32", "=SUM(C28:C31)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCT, align=RIGHT)
    C(ws, "D32", None, fill=FILL_TOTAL)
    C(ws, "E32", None, fill=FILL_TOTAL)
    C(ws, "F32", "=SUM(F28:F31)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "G32", "=SUM(G28:G31)", font=F_LABELB, fill=FILL_TOTAL, fmt=KWD, align=RIGHT)
    C(ws, "H32", "=SUM(H28:H31)", font=F_LABELB, fill=FILL_TOTAL, fmt=PCTc, align=RIGHT)
    add_table(ws, "tblPerfC", "B27:H31")

    C(ws, "B34", "Ending Value (KWD)", font=F_LABELB, align=LEFT)
    C(ws, "C34", "=SUM(G28:G31)", font=F_KPI, fmt=KWD, align=RIGHT)
    C(ws, "B35", "Total Return %", font=F_LABELB, align=LEFT)
    C(ws, "C35", "=IFERROR(C34/C24-1,0)", font=F_KPI, fmt=PCTc, align=RIGHT)
    C(ws, "B36", "CAGR", font=F_LABELB, align=LEFT)
    C(ws, "C36", '=IFERROR((C34/C24)^(1/C25)-1,"")', font=F_KPI, fmt=PCTc, align=RIGHT)

    # SECTION D
    section(ws, "B38:H38", "D.  Growth Comparison & Projection")
    C(ws, "B39", "Projection Horizon (years)", font=F_LABEL, align=LEFT)
    inp(ws, "C39", 10, fmt="0", align=RIGHT)
    note(ws, "B40:H40", "Projections compound the Expected Annual Returns from Settings. "
                        "Portfolio uses the strategic target weights.")

    headers(ws, 42, 2, ["Base Amount (KWD)", "Gold", "Swiss Funds",
                        "S&P 500 ETFs", "US Gov Bonds", "Portfolio"])
    bases = [1, 100000, 1000000]
    for i, b in enumerate(bases):
        r = 43 + i
        C(ws, f"B{r}", b, font=F_LABELB, fmt=KWD, align=RIGHT)
        C(ws, f"C{r}", f"=$B{r}*(1+ExpRet_Gold)^$C$39", fmt=KWD, align=RIGHT)
        C(ws, f"D{r}", f"=$B{r}*(1+ExpRet_Swiss)^$C$39", fmt=KWD, align=RIGHT)
        C(ws, f"E{r}", f"=$B{r}*(1+ExpRet_SP500)^$C$39", fmt=KWD, align=RIGHT)
        C(ws, f"F{r}", f"=$B{r}*(1+ExpRet_Bonds)^$C$39", fmt=KWD, align=RIGHT)
        C(ws, f"G{r}", f"=$B{r}*SUMPRODUCT(Targets,(1+ExpReturns)^$C$39)", fmt=KWD, align=RIGHT)

    # year-by-year projection (chart source) - growth of 1,000,000 KWD
    note(ws, "B47:H47", "Year-by-year growth of 1,000,000 KWD (projection chart source):")
    headers(ws, 48, 2, ["Year", "Gold", "Swiss Funds", "S&P 500 ETFs",
                        "US Gov Bonds", "Portfolio"])
    for y in range(0, 11):
        r = 49 + y
        C(ws, f"B{r}", y, font=F_DATA, fmt="0", align=CENTER)
        C(ws, f"C{r}", f"=1000000*(1+ExpRet_Gold)^$B{r}", fmt=KWD, align=RIGHT)
        C(ws, f"D{r}", f"=1000000*(1+ExpRet_Swiss)^$B{r}", fmt=KWD, align=RIGHT)
        C(ws, f"E{r}", f"=1000000*(1+ExpRet_SP500)^$B{r}", fmt=KWD, align=RIGHT)
        C(ws, f"F{r}", f"=1000000*(1+ExpRet_Bonds)^$B{r}", fmt=KWD, align=RIGHT)
        C(ws, f"G{r}", f"=1000000*SUMPRODUCT(Targets,(1+ExpReturns)^$B{r})", fmt=KWD, align=RIGHT)

    # Charts
    line = LineChart()
    line.title = "Growth Comparison (1,000,000 KWD)"
    line.height = 8
    line.width = 15
    line.add_data(Reference(ws, min_col=3, max_col=7, min_row=48, max_row=59),
                  titles_from_data=True)
    line.set_categories(Reference(ws, min_col=2, min_row=49, max_row=59))
    line.x_axis.title = "Year"
    line.y_axis.title = "Value (KWD)"
    line.x_axis.delete = False
    line.y_axis.delete = False
    ws.add_chart(line, "I5")

    pline = LineChart()
    pline.title = "Portfolio Growth (1,000,000 KWD)"
    pline.height = 8
    pline.width = 15
    pline.add_data(Reference(ws, min_col=7, min_row=48, max_row=59), titles_from_data=True)
    pline.set_categories(Reference(ws, min_col=2, min_row=49, max_row=59))
    pline.x_axis.title = "Year"
    pline.y_axis.title = "Value (KWD)"
    pline.x_axis.delete = False
    pline.y_axis.delete = False
    ws.add_chart(pline, "I22")

    cbar = BarChart()
    cbar.type = "col"
    cbar.title = "Scenario Ending Value by Asset Class (KWD)"
    cbar.height = 8
    cbar.width = 15
    cbar.add_data(Reference(ws, min_col=7, min_row=27, max_row=31), titles_from_data=True)
    cbar.set_categories(Reference(ws, min_col=2, min_row=28, max_row=31))
    cbar.y_axis.numFmt = "#,##0"
    cbar.y_axis.delete = False
    cbar.x_axis.delete = False
    cbar.legend = None
    ws.add_chart(cbar, "I39")

    widths(ws, {"B": 24, "C": 16, "D": 18, "E": 16, "F": 18, "G": 16, "H": 18})
    finish(ws, "A4")


# ============================================================ run builders
build_settings(ws1)
build_inputs(ws2)
build_dashboard(ws3)
build_deployment(ws4)
build_rebalance(ws5)
build_performance(ws6)

# ============================================================ defined names
addname(wb, "Base_Currency", S1, "$C$5")
addname(wb, "Rate_USD", S1, "$C$10")
addname(wb, "Rate_CHF", S1, "$C$11")
addname(wb, "FX_Table", S1, "$B$10:$C$14")
addname(wb, "Target_Gold", S1, "$C$18")
addname(wb, "Target_Swiss", S1, "$C$19")
addname(wb, "Target_SP500", S1, "$C$20")
addname(wb, "Target_Bonds", S1, "$C$21")
addname(wb, "Target_Total", S1, "$C$22")
addname(wb, "Targets", S1, "$C$18:$C$21")
addname(wb, "Tol_Band", S1, "$C$26")
addname(wb, "Min_Action", S1, "$C$30")
addname(wb, "ExpRet_Gold", S1, "$C$36")
addname(wb, "ExpRet_Swiss", S1, "$C$37")
addname(wb, "ExpRet_SP500", S1, "$C$38")
addname(wb, "ExpRet_Bonds", S1, "$C$39")
addname(wb, "ExpReturns", S1, "$C$36:$C$39")

addname(wb, "MV_Gold", S2, "$F$7")
addname(wb, "MV_Swiss", S2, "$F$8")
addname(wb, "MV_SP500", S2, "$F$9")
addname(wb, "MV_Bonds", S2, "$F$10")
addname(wb, "MV_KWD", S2, "$F$7:$F$10")
addname(wb, "Portfolio_Total", S2, "$F$11")

addname(wb, "New_Capital", S4, "$C$3")
addname(wb, "New_Total", S4, "$I$3")

addname(wb, "Val_Date", S6, "$C$3")

wb.save(OUTFILE)
print("Saved", OUTFILE)
