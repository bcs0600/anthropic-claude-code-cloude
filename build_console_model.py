#!/usr/bin/env python3
"""
Swerdlow Multi-Phase Asset Management & Underwriting Platform.

Reshaped per the user's Example.xlsx + Management_Platform Doc:
  * ALL inputs/assumptions live on ONE tab (the Management Console / "Cockpit").
  * Single PROJECT selector (no phase toggle).  Phase No. is a stored field.
  * Side-by-side layout: blue [INPUT] block (cols B-E) beside green [Current]
    block (cols G-J) that mirrors the latest published values from the database.
  * A Publish button writes the blue fields into the relational database as a
    new "Current" record (prior Current -> Historical), which refreshes green.

Delivers .xlsx + importable VBA .bas modules (Excel cannot embed compiled VBA
from pure Python; see README for the one-time import that makes it .xlsm).
"""

import os
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

# ----- conventions ---------------------------------------------------------- #
FONT = "Arial"
C_INPUT = "0000FF"; F_INPUT = "E6F1FB"      # blue input
C_CURR  = "008000"; F_CURR  = "E2EFDA"      # green current
C_HDR   = "1F3864"; F_HDR   = "1F3864"; C_HDRTXT = "FFFFFF"
F_SUB   = "D9E1F2"; F_KPI = "FFF2CC"; F_FLAG = "FFC7CE"; C_NOTE = "806000"

FMT_CCY = '$#,##0;($#,##0);"-"'
FMT_PUPM = '$#,##0.00;($#,##0.00);"-"'
FMT_PCT = '0.0%'; FMT_PCT2 = '0.00%'
FMT_CNT = '#,##0;(#,##0);"-"'
FMT_NUM2 = '0.00'
FMT_DATE = 'mm/dd/yyyy'
FMT_TXT = '@'

fill_input = PatternFill("solid", fgColor=F_INPUT)
fill_curr  = PatternFill("solid", fgColor=F_CURR)
fill_hdr   = PatternFill("solid", fgColor=F_HDR)
fill_sub   = PatternFill("solid", fgColor=F_SUB)
fill_kpi   = PatternFill("solid", fgColor=F_KPI)
fill_flag  = PatternFill("solid", fgColor=F_FLAG)
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
LOCKED = Protection(locked=True); UNLOCKED = Protection(locked=False)

wb = Workbook()

# field registry: each = dict(col=dbcolname, cell=blue A1 on Cockpit)
FIELDS = []
META = ["RecordID", "Project", "Status", "Timestamp", "UpdatedBy", "Notes"]
def reg(name, blue_a1):
    idx = len(FIELDS)
    dbcol = get_column_letter(len(META) + 1 + idx)
    FIELDS.append({"col": name, "cell": blue_a1, "dbcol": dbcol})
    return dbcol

# ----- style helpers -------------------------------------------------------- #
def st(c, *, color="000000", bold=False, size=10, fill=None, fmt=None, align=None,
       wrap=False, border=None, locked=True, italic=False):
    c.font = Font(name=FONT, color=color, bold=bold, size=size, italic=italic)
    if fill is not None: c.fill = fill
    if fmt is not None: c.number_format = fmt
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border is not None: c.border = border
    c.protection = UNLOCKED if not locked else LOCKED
    return c

def lab(ws, a1, text, *, bold=False, indent=0, size=10, italic=False, color="000000"):
    c = ws[a1]; c.value = text
    c.font = Font(name=FONT, color=color, bold=bold, size=size, italic=italic)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=indent)
    return c

def inp(ws, a1, val, *, fmt=FMT_CCY, align="right"):
    c = ws[a1]
    if val is not None: c.value = val
    st(c, color=C_INPUT, fill=fill_input, fmt=fmt, align=align, border=box, locked=False)
    return c

def cur(ws, a1, formula, *, fmt=FMT_CCY, align="right"):
    c = ws[a1]; c.value = formula
    st(c, color=C_CURR, fill=fill_curr, fmt=fmt, align=align, border=box)
    return c

def calc(ws, a1, formula, *, fmt=FMT_CCY, align="right", bold=False, color="000000"):
    c = ws[a1]; c.value = formula
    st(c, color=color, fmt=fmt, align=align, bold=bold, border=box)
    return c

def name(n, sheet, ref):
    wb.defined_names.add(DefinedName(n, attr_text=f"'{sheet}'!{ref}"))

# =========================================================================== #
# LISTS
# =========================================================================== #
def build_lists():
    ws = wb.create_sheet("Lists")
    cols = {
        "A": ("Project", ["RAD 1", "RAD 2", "Live Local A"]),
        "B": ("Status", ["Current", "Historical"]),
        "C": ("UnitType", ["Studio", "1 Bedroom", "2 Bedroom", "3 Bedroom"]),
        "D": ("RentBasis", ["Market", "Live Local", "LIHTC", "PBV", "RAD"]),
        "E": ("YesNo", ["Yes", "No"]),
        "F": ("RentType", ["NNN", "Gross"]),
        "G": ("CostCurve", ["S-curve", "Straight", "Front-loaded", "Back-loaded"]),
    }
    for col, (title, vals) in cols.items():
        st(ws[f"{col}1"], bold=True, color=C_HDR); ws[f"{col}1"] = title
        for i, v in enumerate(vals, start=2):
            ws[f"{col}{i}"] = v; ws[f"{col}{i}"].font = Font(name=FONT, size=10)
    name("Project_List", "Lists", "$A$2:$A$101")
    name("Status_List", "Lists", "$B$2:$B$3")
    name("UnitType_List", "Lists", "$C$2:$C$5")
    name("RentBasis_List", "Lists", "$D$2:$D$6")
    name("YesNo_List", "Lists", "$E$2:$E$3")
    name("RentType_List", "Lists", "$F$2:$F$3")
    name("CostCurve_List", "Lists", "$G$2:$G$5")
    ws.sheet_state = "hidden"; ws.protection.sheet = True

# =========================================================================== #
# MASTER ASSUMPTIONS (program-wide growth defaults; kept as reference driver)
# =========================================================================== #
def build_master():
    ws = wb.create_sheet("Master Assumptions")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    for c in "BCDEF": ws.column_dimensions[c].width = 12
    lab(ws, "A1", "MASTER ASSUMPTIONS  -  program-wide growth defaults (Exhibit E)",
        bold=True, size=13, color=C_HDR)
    lab(ws, "A2", "Per-phase overrides live on the Console. Year 5 = Year 5+.",
        italic=True, color=C_NOTE)
    rows = ["Market Rents", "Affordable Rents", "Retail Rents", "Other Income",
            "Operating Expenses"]
    for j in range(5):
        st(ws.cell(row=4, column=2 + j, value=f"Year {j+1}"), bold=True, fill=fill_sub,
           align="center", border=box)
    st(ws.cell(row=4, column=1, value="Category"), bold=True, fill=fill_sub, border=box)
    defaults = [0.030, 0.025, 0.025, 0.025, 0.030]
    for i, rn in enumerate(rows):
        r = 5 + i
        st(ws.cell(row=r, column=1, value=rn), border=box, align="left")
        for j in range(5):
            inp(ws, f"{get_column_letter(2+j)}{r}", defaults[i], fmt=FMT_PCT)
    name("Growth_Defaults", "Master Assumptions", f"$B$5:$F${5+len(rows)-1}")
    name("Growth_Rows", "Master Assumptions", f"$A$5:$A${5+len(rows)-1}")
    ws.protection.sheet = True

# =========================================================================== #
# COCKPIT  (Management Console) - single tab, side-by-side Input | Current
# =========================================================================== #
def build_cockpit():
    ws = wb.create_sheet("Cockpit")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    for c in "CDE": ws.column_dimensions[c].width = 13
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["G"].width = 26
    for c in "HIJ": ws.column_dimensions[c].width = 13
    ws.column_dimensions["K"].width = 16

    # latest-current-record lookup expression
    maxrec = "MAXIFS(Database!$A:$A,Database!$B:$B,SelectedProject)"
    def green(dbcol, fmt):
        return f'=IFERROR(INDEX(Database!${dbcol}:${dbcol},MATCH({maxrec},Database!$A:$A,0)),"-")'

    # ---------- control bar ----------
    c = ws["B2"]; c.value = "PUBLISH ASSUMPTIONS"
    st(c, color="FFFFFF", bold=True, size=12, fill=PatternFill("solid", fgColor="C00000"),
       align="center", border=box)
    ws.merge_cells("B2:C2")
    lab(ws, "D2", "<- assign Publish_Project macro (placeholder button)", italic=True,
        color=C_NOTE, size=8)
    for a1, txt, col in [("H2", "Add Project", "2E75B6"), ("J2", "Request Support", "2E75B6")]:
        cc = ws[a1]; cc.value = txt
        st(cc, color="FFFFFF", bold=True, fill=PatternFill("solid", fgColor=col),
           align="center", border=box)

    lab(ws, "B4", "MANAGEMENT CONSOLE", bold=True, size=14, color=C_HDR)
    lab(ws, "B5", "Project", bold=True)
    sp = inp(ws, "C5", "RAD 1", fmt=FMT_TXT, align="left")
    name("SelectedProject", "Cockpit", "$C$5")
    dv = DataValidation(type="list", formula1="=Project_List", allow_blank=False)
    ws.add_data_validation(dv); dv.add(ws["C5"])
    lab(ws, "G4", "CURRENT (latest published in database)", bold=True, color=C_CURR)
    lab(ws, "G5", "Last published by"); cur(ws, "H5", f'=IFERROR(INDEX(Database!$E:$E,MATCH({maxrec},Database!$A:$A,0)),"(none)")', fmt=FMT_TXT, align="left")
    lab(ws, "B6", "Record #"); cur(ws, "C6", f"=IFERROR({maxrec},0)", fmt=FMT_CNT, align="left")
    lab(ws, "G6", "Published on"); cur(ws, "H6", f'=IFERROR(INDEX(Database!$D:$D,MATCH({maxrec},Database!$A:$A,0)),"-")', fmt=FMT_DATE, align="left")
    lab(ws, "B7", "Notes (saved on Publish)")
    inp(ws, "C7", "", fmt=FMT_TXT, align="left"); ws.merge_cells("C7:E7")
    name("Console_Notes", "Cockpit", "$C$7")

    # ---------- KPI panel (two bands: Returns & Value, Cost & Physical) ----------
    kpi_cols = ["B", "C", "D", "E", "G", "H", "I", "J", "K"]  # skip narrow spacer F

    def kpi_band(brow, title):
        for cl in "BCDEFGHIJK":
            ws[f"{cl}{brow}"].fill = fill_hdr
        st(ws[f"B{brow}"], color=C_HDRTXT, bold=True, size=10); ws[f"B{brow}"] = title

    def kpi_set(label_row, value_row, items):
        for col, label, formula, fmt in items:
            lc = ws[f"{col}{label_row}"]; lc.value = label
            st(lc, bold=True, fill=fill_kpi, align="center", border=box, size=8, wrap=True)
            vc = ws[f"{col}{value_row}"]; vc.value = formula
            st(vc, bold=True, fill=fill_kpi, align="center", border=box, size=11, fmt=fmt)
        ws.row_dimensions[label_row].height = 24

    kpi_band(9, "KEY METRICS  -  Returns & Value   (from working blue inputs; "
                "IRR / equity multiple / cash-on-cash arrive with the CF engine)")
    kpi_set(10, 11, [
        ("B", "Untrended Yield on Cost", "=IFERROR(NOI_In/TDC_In,0)", FMT_PCT2),
        ("C", "Dev Spread (YoC - Cap)", "=IFERROR(NOI_In/TDC_In,0)-ExitCap_Cell", FMT_PCT2),
        ("D", "Exit Cap Rate", "=ExitCap_Cell", FMT_PCT2),
        ("E", "Stabilized Value", "=IFERROR(NOI_In/ExitCap_Cell,0)", FMT_CCY),
        ("G", "Development Profit", "=IFERROR(NOI_In/ExitCap_Cell,0)-TDC_In", FMT_CCY),
        ("H", "Development Margin", "=IFERROR((NOI_In/ExitCap_Cell-TDC_In)/TDC_In,0)", FMT_PCT),
        ("I", "NOI", "=NOI_In", FMT_CCY),
        ("J", "EGR", "=EGR_In", FMT_CCY),
        ("K", "OpEx Ratio", "=OpExRatio_In", FMT_PCT),
    ])
    kpi_band(12, "KEY METRICS  -  Cost & Physical")
    kpi_set(13, 14, [
        ("B", "Total Dev Cost", "=TDC_In", FMT_CCY),
        ("C", "Cost / Unit", "=IFERROR(TDC_In/Units_Total,0)", FMT_CCY),
        ("D", "Cost / GSF", "=IFERROR(TDC_In/N(GSF_Cell),0)", FMT_PUPM),
        ("E", "Value / Unit", "=IFERROR((NOI_In/ExitCap_Cell)/Units_Total,0)", FMT_CCY),
        ("G", "Total Units", "=Units_Total", FMT_CNT),
        ("H", "Units / Acre", "=IFERROR(Units_Total/N(LandAcres_Cell),0)", FMT_NUM2),
        ("I", "FAR", "=IFERROR(N(GSF_Cell)/(N(LandAcres_Cell)*43560),0)", FMT_NUM2),
        ("J", "Efficiency (NSF/GSF)", "=IFERROR(N(NSF_Cell)/N(GSF_Cell),0)", FMT_PCT),
        ("K", "Avg Unit SF", "=AvgSF_In", FMT_CNT),
    ])

    groups = []
    row = 16

    def section(title):
        nonlocal row
        st(ws.cell(row=row, column=2, value=title), bold=True, color=C_HDRTXT, fill=fill_hdr, size=11)
        for col in range(2, 11):
            ws.cell(row=row, column=col).fill = fill_hdr
        st(ws.cell(row=row, column=7, value="Input  |  Current ->"), color=C_HDRTXT, size=8, align="right")
        row += 1
        return row - 1

    def subhdr(in_cols, cur_cols):
        """in_cols/cur_cols: list of (col_letter, title)."""
        nonlocal row
        for col, t in in_cols:
            st(ws[f"{col}{row}"], bold=True, fill=fill_sub, align="center", border=box, size=8)
            ws[f"{col}{row}"] = t
        for col, t in cur_cols:
            st(ws[f"{col}{row}"], bold=True, fill=fill_sub, align="center", border=box, size=8)
            ws[f"{col}{row}"] = t
        row += 1

    def scalar(label, colname, default, fmt, *, validation=None, required=False, in_col="D"):
        """One assumption: label in B, blue input in in_col, green current in mirror col."""
        nonlocal row
        lab(ws, f"B{row}", label, indent=1)
        lab(ws, f"G{row}", label, indent=1, color=C_CURR)
        inp(ws, f"{in_col}{row}", default, fmt=fmt)
        if validation:
            dvv = DataValidation(type="list", formula1=f"={validation}", allow_blank=True)
            ws.add_data_validation(dvv); dvv.add(ws[f"{in_col}{row}"])
        dbcol = reg(colname, f"Cockpit!${in_col}${row}")
        cur_col = {"C": "H", "D": "I", "E": "J"}[in_col]
        cur(ws, f"{cur_col}{row}", green(dbcol, fmt), fmt=fmt)
        if required:
            st(ws[f"F{row}"], color="C00000", align="center", size=9)
            ws[f"F{row}"] = f'=IF(N({in_col}{row})=0,"!","")' if fmt != FMT_TXT else f'=IF({in_col}{row}="","!","")'
            ws.conditional_formatting.add(f"{in_col}{row}",
                FormulaRule(formula=[f'ISBLANK({in_col}{row})'], fill=fill_flag))
        r = row; row += 1
        return r

    # ======================= (1) General Inputs =======================
    s = section("General Inputs")
    scalar("Phase No.", "PhaseNo", 3, FMT_CNT)
    scalar("Project Name", "ProjectName", "RAD 1", FMT_TXT, required=True)
    scalar("Land Size (Acres)", "LandAcres", 2, FMT_NUM2)
    name("LandAcres_Cell", "Cockpit", f"$D${row-1}")
    scalar("Closing Date", "ClosingDate", date(2026, 7, 1), FMT_DATE, required=True)
    name("Closing_Date_Cell", "Cockpit", f"$D${row-1}")
    scalar("Gross Building Area (GSF)", "GSF", 240000, FMT_CNT, required=True)
    name("GSF_Cell", "Cockpit", f"$D${row-1}")
    scalar("Net Rentable Area (NSF)", "NSF", 200000, FMT_CNT)
    name("NSF_Cell", "Cockpit", f"$D${row-1}")
    scalar("Exit / Stabilized Cap Rate", "ExitCapRate", 0.055, FMT_PCT)
    name("ExitCap_Cell", "Cockpit", f"$D${row-1}")
    groups.append((s + 1, row - 1))

    # ======================= (2) Key Dates =======================
    s = section("Key Dates")
    for nm, key, dv in [("Construction Start", "off_ConstStart", 2),
                        ("Initial Occupancy", "off_InitOcc", 14),
                        ("Construction End", "off_ConstEnd", 20),
                        ("Stabilization", "off_Stab", 30)]:
        lab(ws, f"B{row}", f"{nm} (month offset)", indent=1)
        lab(ws, f"G{row}", nm, indent=1, color=C_CURR)
        inp(ws, f"D{row}", dv, fmt=FMT_CNT)
        dbcol = reg(key, f"Cockpit!$D${row}")
        calc(ws, f"E{row}", f"=EDATE(Closing_Date_Cell,D{row})", fmt=FMT_DATE)
        cur(ws, f"I{row}", green(dbcol, FMT_CNT), fmt=FMT_CNT)
        cur(ws, f"J{row}", f"=IFERROR(EDATE(Closing_Date_Cell,I{row}),\"-\")", fmt=FMT_DATE)
        row += 1
    groups.append((s + 1, row - 1))

    # ======================= (3) Commercial Assumptions =======================
    s = section("Commercial Assumptions")
    scalar("Has Commercial?", "HasCommercial", "No", FMT_TXT, validation="YesNo_List")
    name("HasComm_Cell", "Cockpit", f"$D${row-1}")
    scalar("Commercial Leasable SF", "CommSF", 0, FMT_CNT)
    name("CommSF_Cell", "Cockpit", f"$D${row-1}")
    scalar("Annual Rent PSF", "CommRentPSF", 30, FMT_PUPM)
    name("CommPSF_Cell", "Cockpit", f"$D${row-1}")
    scalar("Escalation", "CommEscalation", 0.03, FMT_PCT)
    scalar("Abatement (months)", "CommAbatement", 0, FMT_CNT)
    scalar("General Vacancy %", "CommVacancy", 0.07, FMT_PCT)
    name("CommVac_Cell", "Cockpit", f"$D${row-1}")
    scalar("Operating Expenses $/SF", "CommOpExPSF", 6.75, FMT_PUPM)
    name("CommOpEx_Cell", "Cockpit", f"$D${row-1}")
    # commercial NOI (both sides)
    lab(ws, f"B{row}", "Stabilized Commercial NOI", bold=True, indent=1)
    lab(ws, f"G{row}", "Stabilized Commercial NOI", bold=True, indent=1, color=C_CURR)
    calc(ws, f"E{row}", "=IF(HasComm_Cell=\"Yes\",CommSF_Cell*CommPSF_Cell*(1-CommVac_Cell)-CommSF_Cell*CommOpEx_Cell,0)", fmt=FMT_CCY, bold=True)
    name("CommNOI_In", "Cockpit", f"$E${row}")
    # green commercial NOI from green cells
    gc = {k["col"]: k["dbcol"] for k in FIELDS}
    calc(ws, f"J{row}", f'=IFERROR(IF(INDEX(Database!${gc["HasCommercial"]}:${gc["HasCommercial"]},MATCH({maxrec},Database!$A:$A,0))="Yes",'
                        f'INDEX(Database!${gc["CommSF"]}:${gc["CommSF"]},MATCH({maxrec},Database!$A:$A,0))*'
                        f'INDEX(Database!${gc["CommRentPSF"]}:${gc["CommRentPSF"]},MATCH({maxrec},Database!$A:$A,0))*'
                        f'(1-INDEX(Database!${gc["CommVacancy"]}:${gc["CommVacancy"]},MATCH({maxrec},Database!$A:$A,0)))-'
                        f'INDEX(Database!${gc["CommSF"]}:${gc["CommSF"]},MATCH({maxrec},Database!$A:$A,0))*'
                        f'INDEX(Database!${gc["CommOpExPSF"]}:${gc["CommOpExPSF"]},MATCH({maxrec},Database!$A:$A,0)),0),0)',
         fmt=FMT_CCY, bold=True, color=C_CURR)
    name("CommNOI_Cur", "Cockpit", f"$J${row}")
    row += 1
    groups.append((s + 1, row - 1))

    # ======================= (4) Multifamily Unit Mix + Revenue =======================
    s = section("Multifamily - Unit Mix & Rental Revenue (Exhibit A/C)")
    subhdr([("B", "Unit Type"), ("C", "Count"), ("D", "Avg SF"), ("E", "$ PUPM")],
           [("G", "Unit Type"), ("H", "Count"), ("I", "Avg SF"), ("J", "$ PUPM")])
    units = ["Studio", "1 Bedroom", "2 Bedroom", "3 Bedroom"]
    um_first = row
    for i, u in enumerate(units):
        r = um_first + i
        st(ws[f"B{r}"], color=C_INPUT, fill=fill_input, align="left", border=box, locked=False)
        ws[f"B{r}"] = u
        dvb = DataValidation(type="list", formula1="=UnitType_List", allow_blank=True)
        ws.add_data_validation(dvb); dvb.add(ws[f"B{r}"])
        reg(f"UM{i+1}_Type", f"Cockpit!$B${r}")
        inp(ws, f"C{r}", [40, 80, 60, 20][i], fmt=FMT_CNT); reg(f"UM{i+1}_Count", f"Cockpit!$C${r}")
        inp(ws, f"D{r}", [550, 750, 1050, 1300][i], fmt=FMT_CNT); reg(f"UM{i+1}_AvgSF", f"Cockpit!$D${r}")
        inp(ws, f"E{r}", [1800, 2200, 2900, 3500][i], fmt=FMT_CCY); reg(f"UM{i+1}_PUPM", f"Cockpit!$E${r}")
        # green mirrors
        lab(ws, f"G{r}", u, color=C_CURR, indent=0)
        cur(ws, f"H{r}", green(gc2(f"UM{i+1}_Count"), FMT_CNT), fmt=FMT_CNT)
        cur(ws, f"I{r}", green(gc2(f"UM{i+1}_AvgSF"), FMT_CNT), fmt=FMT_CNT)
        cur(ws, f"J{r}", green(gc2(f"UM{i+1}_PUPM"), FMT_CCY), fmt=FMT_CCY)
    um_last = um_first + len(units) - 1
    # totals row
    tr = um_last + 1
    st(ws.cell(row=tr, column=2, value="Total / Blended"), bold=True, fill=fill_sub, border=box, align="left")
    st(ws.cell(row=tr, column=7, value="Total / Blended"), bold=True, fill=fill_sub, border=box, align="left", color=C_CURR)
    calc(ws, f"C{tr}", f"=SUM(C{um_first}:C{um_last})", fmt=FMT_CNT, bold=True)
    calc(ws, f"D{tr}", f"=IFERROR(SUMPRODUCT(D{um_first}:D{um_last},C{um_first}:C{um_last})/C{tr},0)", fmt=FMT_CNT, bold=True)
    calc(ws, f"E{tr}", f"=IFERROR(SUMPRODUCT(E{um_first}:E{um_last},C{um_first}:C{um_last})/C{tr},0)", fmt=FMT_CCY, bold=True)
    calc(ws, f"H{tr}", f"=SUM(H{um_first}:H{um_last})", fmt=FMT_CNT, bold=True, color=C_CURR)
    calc(ws, f"I{tr}", f"=IFERROR(SUMPRODUCT(I{um_first}:I{um_last},H{um_first}:H{um_last})/H{tr},0)", fmt=FMT_CNT, bold=True, color=C_CURR)
    calc(ws, f"J{tr}", f"=IFERROR(SUMPRODUCT(J{um_first}:J{um_last},H{um_first}:H{um_last})/H{tr},0)", fmt=FMT_CCY, bold=True, color=C_CURR)
    name("Units_Total", "Cockpit", f"$C${tr}")
    name("Units_Total_Cur", "Cockpit", f"$H${tr}")
    name("Blended_PUPM", "Cockpit", f"$E${tr}")
    name("Blended_PUPM_Cur", "Cockpit", f"$J${tr}")
    name("AvgSF_In", "Cockpit", f"$D${tr}")
    row = tr + 2

    # revenue waterfall (GPR -> Effective Rental Revenue), both sides
    def wf(label, in_formula, cur_formula, fmt=FMT_CCY, bold=False, in_drv=None, cur_drv=None, drvfmt=FMT_PCT, key=None):
        nonlocal row
        lab(ws, f"B{row}", label, indent=1, bold=bold)
        lab(ws, f"G{row}", label, indent=1, bold=bold, color=C_CURR)
        if in_drv is not None:
            inp(ws, f"D{row}", in_drv, fmt=drvfmt)
            dbcol = reg(key, f"Cockpit!$D${row}")
            cur(ws, f"I{row}", green(dbcol, drvfmt), fmt=drvfmt)
        calc(ws, f"E{row}", in_formula, fmt=fmt, bold=bold)
        calc(ws, f"J{row}", cur_formula, fmt=fmt, bold=bold, color=C_CURR)
        r = row; row += 1; return r
    gpr = wf("Gross Potential Rent", f"=Blended_PUPM*Units_Total*12", f"=Blended_PUPM_Cur*Units_Total_Cur*12", bold=True)
    name("GPR_In", "Cockpit", f"$E${gpr}"); name("GPR_Cur", "Cockpit", f"$J${gpr}")
    vac = wf("Vacancy", f"=-GPR_In*D{row}", f"=-GPR_Cur*I{row}", in_drv=0.05, key="ResVacancy")
    con = wf("Concessions", f"=-GPR_In*D{row}", f"=-GPR_Cur*I{row}", in_drv=0.0, key="ResConcessions")
    mod = wf("Model", f"=-GPR_In*D{row}", f"=-GPR_Cur*I{row}", in_drv=0.0, key="ResModel")
    col_ = wf("Collection Loss", f"=-GPR_In*D{row}", f"=-GPR_Cur*I{row}", in_drv=0.0, key="ResCollLoss")
    err = wf("Effective Rental Revenue", f"=GPR_In+E{vac}+E{con}+E{mod}+E{col_}",
             f"=GPR_Cur+J{vac}+J{con}+J{mod}+J{col_}", bold=True)
    name("ERR_In", "Cockpit", f"$E${err}"); name("ERR_Cur", "Cockpit", f"$J${err}")
    groups.append((s + 1, row - 1))

    # ======================= (5) Other Income =======================
    s = section("Other Residential Income (Exhibit C)")
    subhdr([("B", "Item"), ("D", "$ PUPM"), ("E", "$ Annual")],
           [("G", "Item"), ("I", "$ PUPM"), ("J", "$ Annual")])
    oi_first = row
    for i in range(5):
        r = oi_first + i
        st(ws[f"B{r}"], color=C_INPUT, fill=fill_input, align="left", border=box, locked=False)
        ws[f"B{r}"] = f"Item {i+1}"
        reg(f"OI{i+1}_Name", f"Cockpit!$B${r}")
        inp(ws, f"D{r}", 100 if i == 0 else "", fmt=FMT_PUPM); reg(f"OI{i+1}_PUPM", f"Cockpit!$D${r}")
        calc(ws, f"E{r}", f"=N(D{r})*12*Units_Total", fmt=FMT_CCY)
        lab(ws, f"G{r}", f"Item {i+1}", color=C_CURR)
        cur(ws, f"I{r}", green(gc2(f"OI{i+1}_PUPM"), FMT_PUPM), fmt=FMT_PUPM)
        cur(ws, f"J{r}", f"=N(I{r})*12*Units_Total_Cur", fmt=FMT_CCY)
    oi_last = oi_first + 4
    tr = oi_last + 1
    st(ws.cell(row=tr, column=2, value="Total Other Income"), bold=True, fill=fill_sub, border=box, align="left")
    st(ws.cell(row=tr, column=7, value="Total Other Income"), bold=True, fill=fill_sub, border=box, align="left", color=C_CURR)
    calc(ws, f"E{tr}", f"=SUM(E{oi_first}:E{oi_last})", fmt=FMT_CCY, bold=True)
    calc(ws, f"J{tr}", f"=SUM(J{oi_first}:J{oi_last})", fmt=FMT_CCY, bold=True, color=C_CURR)
    name("OtherInc_In", "Cockpit", f"$E${tr}"); name("OtherInc_Cur", "Cockpit", f"$J${tr}")
    row = tr + 2
    # EGR
    egr = row
    st(ws.cell(row=egr, column=2, value="Effective Gross Revenue (EGR)"), bold=True, fill=fill_sub, border=box, align="left")
    st(ws.cell(row=egr, column=7, value="Effective Gross Revenue (EGR)"), bold=True, fill=fill_sub, border=box, align="left", color=C_CURR)
    calc(ws, f"E{egr}", "=ERR_In+OtherInc_In", fmt=FMT_CCY, bold=True)
    calc(ws, f"J{egr}", "=ERR_Cur+OtherInc_Cur", fmt=FMT_CCY, bold=True, color=C_CURR)
    name("EGR_In", "Cockpit", f"$E${egr}"); name("EGR_Cur", "Cockpit", f"$J${egr}")
    row = egr + 2
    groups.append((s + 1, row - 1))

    # ======================= (6) Operating Expenses =======================
    s = section("Multifamily Operating Expenses (Exhibit D)")
    subhdr([("B", "Item"), ("D", "$ PUPA"), ("E", "$ Annual")],
           [("G", "Item"), ("I", "$ PUPA"), ("J", "$ Annual")])
    opex = [("Payroll", "OpEx_Payroll", 1150), ("Marketing", "OpEx_Marketing", 350),
            ("Administrative", "OpEx_Admin", 350), ("Professional Fees", "OpEx_ProfFees", 10),
            ("Repairs & Maint.", "OpEx_RM", 300), ("Utilities", "OpEx_Utilities", 900),
            ("Insurance", "OpEx_Insurance", 1000), ("RE Taxes", "OpEx_RETaxes", 4500)]
    ox_first = row
    ox_total_in, ox_total_cur = [], []
    for nm, key, dv in opex:
        r = row
        lab(ws, f"B{r}", nm, indent=1); lab(ws, f"G{r}", nm, indent=1, color=C_CURR)
        inp(ws, f"D{r}", dv, fmt=FMT_CCY); dbcol = reg(key, f"Cockpit!$D${r}")
        calc(ws, f"E{r}", f"=N(D{r})*Units_Total", fmt=FMT_CCY); ox_total_in.append(f"E{r}")
        cur(ws, f"I{r}", green(dbcol, FMT_CCY), fmt=FMT_CCY)
        calc(ws, f"J{r}", f"=N(I{r})*Units_Total_Cur", fmt=FMT_CCY, color=C_CURR); ox_total_cur.append(f"J{r}")
        row += 1
    # Management fee (% of EGR)
    r = row
    lab(ws, f"B{r}", "Management Fee (% of EGR)", indent=1); lab(ws, f"G{r}", "Management Fee (% of EGR)", indent=1, color=C_CURR)
    inp(ws, f"D{r}", 0.03, fmt=FMT_PCT); dbcol = reg("OpEx_MgmtFeePct", f"Cockpit!$D${r}")
    calc(ws, f"E{r}", "=D{0}*EGR_In".format(r), fmt=FMT_CCY); ox_total_in.append(f"E{r}")
    cur(ws, f"I{r}", green(dbcol, FMT_PCT), fmt=FMT_PCT)
    calc(ws, f"J{r}", f"=I{r}*EGR_Cur", fmt=FMT_CCY, color=C_CURR); ox_total_cur.append(f"J{r}")
    row += 1
    # total opex
    tr = row
    st(ws.cell(row=tr, column=2, value="Total Operating Expenses"), bold=True, fill=fill_sub, border=box, align="left")
    st(ws.cell(row=tr, column=7, value="Total Operating Expenses"), bold=True, fill=fill_sub, border=box, align="left", color=C_CURR)
    calc(ws, f"E{tr}", "=" + "+".join(ox_total_in), fmt=FMT_CCY, bold=True)
    calc(ws, f"J{tr}", "=" + "+".join(ox_total_cur), fmt=FMT_CCY, bold=True, color=C_CURR)
    name("TotOpEx_In", "Cockpit", f"$E${tr}"); name("TotOpEx_Cur", "Cockpit", f"$J${tr}")
    row = tr + 2
    groups.append((s + 1, row - 1))

    # ======================= (7) NOI / OpEx Ratio =======================
    s = section("Net Operating Income")
    noi = row
    st(ws.cell(row=noi, column=2, value="Net Operating Income"), bold=True, fill=fill_sub, border=box, align="left")
    st(ws.cell(row=noi, column=7, value="Net Operating Income"), bold=True, fill=fill_sub, border=box, align="left", color=C_CURR)
    calc(ws, f"E{noi}", "=EGR_In-TotOpEx_In+CommNOI_In", fmt=FMT_CCY, bold=True)
    calc(ws, f"J{noi}", "=EGR_Cur-TotOpEx_Cur+CommNOI_Cur", fmt=FMT_CCY, bold=True, color=C_CURR)
    name("NOI_In", "Cockpit", f"$E${noi}"); name("NOI_Cur", "Cockpit", f"$J${noi}")
    row += 1
    oxr = row
    lab(ws, f"B{oxr}", "OpEx Ratio"); lab(ws, f"G{oxr}", "OpEx Ratio", color=C_CURR)
    calc(ws, f"E{oxr}", "=IFERROR(TotOpEx_In/EGR_In,0)", fmt=FMT_PCT)
    calc(ws, f"J{oxr}", "=IFERROR(TotOpEx_Cur/EGR_Cur,0)", fmt=FMT_PCT, color=C_CURR)
    name("OpExRatio_In", "Cockpit", f"$E${oxr}")
    row += 2
    groups.append((s + 1, row - 1))

    # ======================= (8) Development Budget =======================
    s = section("Development Budget / Capital Costs")
    subhdr([("B", "Line"), ("D", "$"), ("E", "$/Unit")],
           [("G", "Line"), ("I", "$"), ("J", "$/Unit")])
    budget = [("Land", "Bud_Land", 8000000), ("Shell / Core", "Bud_Shell", 28000000),
              ("Sitework", "Bud_Sitework", 4000000), ("Parking Structure", "Bud_Parking", 6000000),
              ("FF&E", "Bud_FFE", 1500000), ("GC Fee", "Bud_GCFee", 2000000),
              ("A&E", "Bud_AE", 2500000), ("Permits & Impact Fees", "Bud_Permits", 1800000),
              ("Legal", "Bud_Legal", 600000), ("Marketing / Lease-up", "Bud_Mktg", 700000),
              ("Developer Fee", "Bud_DevFee", 3000000),
              ("Taxes & Insurance (Constr.)", "Bud_TaxIns", 900000)]
    bud_in, bud_cur = [], []
    for nm, key, dv in budget:
        r = row
        lab(ws, f"B{r}", nm, indent=1); lab(ws, f"G{r}", nm, indent=1, color=C_CURR)
        inp(ws, f"D{r}", dv, fmt=FMT_CCY); dbcol = reg(key, f"Cockpit!$D${r}")
        calc(ws, f"E{r}", f"=IFERROR(D{r}/Units_Total,0)", fmt=FMT_CCY); bud_in.append(f"D{r}")
        cur(ws, f"I{r}", green(dbcol, FMT_CCY), fmt=FMT_CCY)
        calc(ws, f"J{r}", f"=IFERROR(I{r}/Units_Total_Cur,0)", fmt=FMT_CCY, color=C_CURR); bud_cur.append(f"I{r}")
        row += 1
    tr = row
    st(ws.cell(row=tr, column=2, value="Total Development Cost"), bold=True, fill=fill_sub, border=box, align="left")
    st(ws.cell(row=tr, column=7, value="Total Development Cost"), bold=True, fill=fill_sub, border=box, align="left", color=C_CURR)
    calc(ws, f"E{tr}", "=" + "+".join(bud_in), fmt=FMT_CCY, bold=True)
    calc(ws, f"J{tr}", "=" + "+".join(bud_cur), fmt=FMT_CCY, bold=True, color=C_CURR)
    name("TDC_In", "Cockpit", f"$E${tr}"); name("TDC_Cur", "Cockpit", f"$J${tr}")
    row = tr + 2
    groups.append((s + 1, row - 1))

    # grouping + protection
    for a, b in groups:
        if b >= a:
            ws.row_dimensions.group(a, b, outline_level=1, hidden=False)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.protection.sheet = True
    return ws

# helper to fetch a dbcol after registration (used inside build_cockpit closures)
def gc2(colname):
    return next(f["dbcol"] for f in FIELDS if f["col"] == colname)

# =========================================================================== #
# DATABASE (relational, versioned)
# =========================================================================== #
def build_database():
    ws = wb.create_sheet("Database")
    ws.sheet_view.showGridLines = False
    lab(ws, "A1", "DATABASE  -  relational system of record (one row = one published record). Do not edit by hand.",
        bold=True, color=C_HDR)
    lab(ws, "A2", "MaxRecordsPerProject"); inp(ws, "B2", 20, fmt=FMT_CNT, align="left")
    name("MaxRecordsPerProject", "Database", "$B$2")
    lab(ws, "C2", "FIFO cap; oldest record per Project is archived on overflow. Status: Current vs Historical.",
        italic=True, color=C_NOTE, size=8)
    hdr = META + [f["col"] for f in FIELDS]
    for j, h in enumerate(hdr):
        st(ws.cell(row=4, column=1 + j, value=h), bold=True, fill=fill_hdr, color=C_HDRTXT,
           align="center", border=box, size=8)
    # seed: one Current record per sample project; first one mirrors live Console inputs
    seeds = [("RAD 1", 1), ("RAD 2", 0), ("Live Local A", 0)]
    rr = 5
    for k, (proj, live) in enumerate(seeds, start=1):
        ws.cell(row=rr, column=1, value=k)
        ws.cell(row=rr, column=2, value=proj)
        ws.cell(row=rr, column=3, value="Current")
        ws.cell(row=rr, column=4, value=date(2026, 6, 23))
        ws.cell(row=rr, column=5, value="seed")
        ws.cell(row=rr, column=6, value="initial seed record")
        for j, fld in enumerate(FIELDS):
            col = len(META) + 1 + j
            if live:
                ws.cell(row=rr, column=col, value=f"={fld['cell']}")
            else:
                ws.cell(row=rr, column=col, value="")
        rr += 1
    for rrow in ws[f"A4:{get_column_letter(len(hdr))}{rr-1}"]:
        for c in rrow:
            if c.row >= 5: c.font = Font(name=FONT, size=8)
    ws.column_dimensions["A"].width = 8
    ws.protection.sheet = True

def build_archive():
    ws = wb.create_sheet("Archive")
    lab(ws, "A1", "ARCHIVE  -  purged records (FIFO overflow). Recoverable, not hard-deleted.",
        bold=True, color=C_HDR)
    hdr = META + [f["col"] for f in FIELDS]
    for j, h in enumerate(hdr):
        st(ws.cell(row=3, column=1 + j, value=h), bold=True, fill=fill_hdr, color=C_HDRTXT, size=8)
    ws.sheet_state = "hidden"; ws.protection.sheet = True

# =========================================================================== #
# PORTFOLIO DASHBOARD (roll-up by Project, read-only)
# =========================================================================== #
def build_portfolio():
    ws = wb.create_sheet("Portfolio Dashboard")
    ws.sheet_view.showGridLines = False
    lab(ws, "A1", "PORTFOLIO DASHBOARD  -  all projects (latest published record), read-only",
        bold=True, size=13, color=C_HDR)
    cols = ["Project", "Phase No.", "Units", "NSF", "EGR", "NOI", "OpEx Ratio",
            "Total Dev Cost", "Cost/Unit", "Yield on Cost"]
    for j, h in enumerate(cols):
        st(ws.cell(row=3, column=1 + j, value=h), bold=True, fill=fill_hdr, color=C_HDRTXT,
           align="center", border=box, size=9)
        ws.column_dimensions[get_column_letter(1 + j)].width = 14
    def dbget(colname, pcell):
        col = gc2(colname)
        return (f'INDEX(Database!${col}:${col},MATCH(MAXIFS(Database!$A:$A,Database!$B:$B,{pcell}),'
                f'Database!$A:$A,0))')
    first = 4
    for i in range(20):
        r = first + i
        pcell = f"Lists!$A${2+i}"
        ws.cell(row=r, column=1, value=f'=IF({pcell}="","",{pcell})')
        ws.cell(row=r, column=2, value=f'=IFERROR(IF({pcell}="","",{dbget("PhaseNo",pcell)}),"")')
        # units = sum of UM counts in the latest record
        rec = f'MATCH(MAXIFS(Database!$A:$A,Database!$B:$B,{pcell}),Database!$A:$A,0)'
        ucols = [gc2(f"UM{n}_Count") for n in range(1, 5)]
        uterm = "+".join([f'N(INDEX(Database!${c}:${c},{rec}))' for c in ucols])
        ws.cell(row=r, column=3, value=f'=IFERROR(IF({pcell}="","",{uterm}),"")')
        ws.cell(row=r, column=4, value=f'=IFERROR(IF({pcell}="","",{dbget("NSF",pcell)}),"")')
        for cc in range(5, 11):
            ws.cell(row=r, column=cc, value=f'=IF({pcell}="","","-")')
        for cc in range(1, 11):
            st(ws.cell(row=r, column=cc), border=box, size=9, color=C_CURR,
               align=("left" if cc == 1 else "right"),
               fmt=(FMT_CNT if cc in (2, 3, 4) else (FMT_TXT if cc == 1 else FMT_CCY)))
    tr = first + 20
    st(ws.cell(row=tr, column=1, value="PORTFOLIO TOTAL"), bold=True, fill=fill_sub, border=box, align="left")
    ws.cell(row=tr, column=3, value=f"=SUM(C{first}:C{first+19})")
    ws.cell(row=tr, column=4, value=f"=SUM(D{first}:D{first+19})")
    for cc in (3, 4):
        st(ws.cell(row=tr, column=cc), bold=True, fill=fill_sub, border=box, fmt=FMT_CNT, align="right")
    lab(ws, f"A{tr+2}", "Note: EGR/NOI/cost $ roll-ups recompute from each project's stored "
        "inputs (DB holds raw assumptions); wired fully alongside the calc engine next pass.",
        italic=True, color=C_NOTE, size=8)
    ws.protection.sheet = True

# =========================================================================== #
# PROJECT DASHBOARD (single project, read-only)
# =========================================================================== #
def build_project_dash():
    ws = wb.create_sheet("Project Dashboard")
    ws.sheet_view.showGridLines = False
    lab(ws, "A1", "PROJECT DASHBOARD  -  read-only single-project view (printable)",
        bold=True, size=13, color=C_HDR)
    lab(ws, "A3", "Project")
    c = inp(ws, "B3", "RAD 1", fmt=FMT_TXT, align="left")
    name("SelectedProjectDash", "Project Dashboard", "$B$3")
    dv = DataValidation(type="list", formula1="=Project_List", allow_blank=False)
    ws.add_data_validation(dv); dv.add(c)
    lab(ws, "A5", "Renders the operating proforma for the selected project from the database "
        "latest record. (Full recompute wired with the calc engine next pass.)",
        italic=True, color=C_NOTE)
    lines = ["Gross Potential Rent", "Effective Rental Revenue", "Other Income",
             "Effective Gross Revenue (EGR)", "Total Operating Expenses", "Net Operating Income",
             "OpEx Ratio"]
    st(ws.cell(row=7, column=1, value="Operating Proforma"), bold=True, fill=fill_hdr, color=C_HDRTXT)
    st(ws.cell(row=7, column=2, value="$ Annual"), bold=True, fill=fill_hdr, color=C_HDRTXT, align="center")
    for i, ln in enumerate(lines):
        r = 8 + i
        lab(ws, f"A{r}", ln, indent=1)
        ws.cell(row=r, column=2, value='="-"')
        st(ws.cell(row=r, column=2), color=C_CURR, align="right",
           fmt=(FMT_PCT if ln == "OpEx Ratio" else FMT_CCY), border=box)
    ws.protection.sheet = True

# =========================================================================== #
# MONTHLY CF (placeholder date spine)
# =========================================================================== #
def build_monthly():
    ws = wb.create_sheet("Monthly CF")
    ws.sheet_view.showGridLines = False
    lab(ws, "A1", "MONTHLY CF  -  FUTURE calc engine placeholder (date spine only this pass)",
        bold=True, size=13, color="C00000")
    lab(ws, "A2", "Spine driven by the Console closing date for the selected project. "
        "Operating CFs / capital costs / unlevered CFs plug in below later.",
        italic=True, color=C_NOTE)
    lab(ws, "A4", "Closing date (from Console)")
    cur(ws, "B4", "=Closing_Date_Cell", fmt=FMT_DATE, align="left")
    lab(ws, "A6", "Month index"); lab(ws, "A7", "Month start")
    for m in range(120):
        ws.cell(row=6, column=3 + m, value=m + 1)
        st(ws.cell(row=6, column=3 + m), align="center", size=8, bold=True, fill=fill_sub)
        ws.cell(row=7, column=3 + m, value=f"=EDATE($B$4,{m})")
        st(ws.cell(row=7, column=3 + m), fmt="mmm-yy", align="center", size=8, color=C_CURR)
    ws.protection.sheet = True

# =========================================================================== #
# READ ME
# =========================================================================== #
def build_readme():
    ws = wb.create_sheet("Read me")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 112
    rows = [
        ("Swerdlow Multi-Phase Asset Management & Underwriting Platform", True, 15, C_HDR),
        ("", False, 9, "000000"),
        ("THE LOOP", True, 12, C_HDR),
        ("Pick a PROJECT (single selector, top of the Console) -> the GREEN 'Current' column "
         "shows the latest assumptions published to the database -> edit the BLUE input fields "
         "-> click PUBLISH. Publish validates, writes the blue values to the database as a new "
         "'Current' record (prior Current -> Historical), and the green values refresh.", False, 10, "000000"),
        ("", False, 9, "000000"),
        ("ONE TAB FOR ALL INPUTS", True, 12, C_HDR),
        ("Every assumption lives on the Cockpit (Management Console) tab, laid out as a blue "
         "[INPUT] block (cols B-E) beside a green [Current] block (cols G-J) that mirrors the "
         "database. Black = calculated. The database is the single source of truth.", False, 10, "000000"),
        ("", False, 9, "000000"),
        ("TABS", True, 12, C_HDR),
        ("Read me | Cockpit (console) | Master Assumptions | Database | Portfolio Dashboard | "
         "Project Dashboard | Monthly CF.  Hidden: Lists, Archive.", False, 10, "000000"),
        ("", False, 9, "000000"),
        ("** ENABLING THE MACROS (.xlsx -> .xlsm) **", True, 12, "C00000"),
        ("Pure-Python tooling cannot author Excel's compiled VBA, so the macros ship as .bas "
         "files in ./vba. To enable:", False, 10, "000000"),
        ("  1. Open in Excel -> Save As -> Excel Macro-Enabled Workbook (.xlsm).", False, 10, "000000"),
        ("  2. Alt+F11 -> File > Import File... -> import each .bas in ./vba.", False, 10, "000000"),
        ("  3. Assign Publish_Project to the red PUBLISH ASSUMPTIONS button, and Add_Project / "
         "Request_Support to their buttons.", False, 10, "000000"),
        ("  4. Set CALENDLY_URL in Module3_Support.", False, 10, "000000"),
        ("", False, 9, "000000"),
        ("VERSION CONTROL", True, 12, C_HDR),
        ("Each Publish appends a record with User, Timestamp, Notes, and Status. Latest = Current; "
         "prior records for that project become Historical. FIFO cap (MaxRecordsPerProject, default "
         "20) archives the oldest record per project to the hidden Archive tab (auditable).", False, 10, "000000"),
    ]
    r = 1
    for text, bold, size, color in rows:
        c = ws[f"A{r}"]; c.value = text
        c.font = Font(name=FONT, bold=bold, size=size, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30 if len(text) > 95 else (18 if text else 8)
        r += 1
    ws.protection.sheet = True

# =========================================================================== #
# BUILD
# =========================================================================== #
build_lists()
build_master()
build_cockpit()
build_database()
build_archive()
build_portfolio()
build_project_dash()
build_monthly()
build_readme()
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]
order = ["Read me", "Cockpit", "Master Assumptions", "Database", "Portfolio Dashboard",
         "Project Dashboard", "Monthly CF", "Lists", "Archive"]
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
wb.calculation.fullCalcOnLoad = True
out = "/home/user/anthropic-claude-code-cloude/Swerdlow_Management_Platform.xlsx"
wb.save(out)
print("Saved", out, "| fields:", len(FIELDS))

# ----- VBA -----------------------------------------------------------------
def write_vba():
    vdir = "/home/user/anthropic-claude-code-cloude/vba_platform"
    os.makedirs(vdir, exist_ok=True)
    projname_addr = next(f["cell"] for f in FIELDS if f["col"] == "ProjectName").replace("Cockpit!", "")
    addr_lines = "\n".join(f'    A({i+1}) = "{f["cell"].replace("Cockpit!","")}"' for i, f in enumerate(FIELDS))
    m1 = f'''Attribute VB_Name = "Module1_Publish"
'================================================================
' Publish_Project  -  writes the blue Console inputs into the
' relational Database as a new "Current" record (prior Current ->
' Historical), then enforces the FIFO version cap.
' Database: row 4 headers, data from row 5.
'   A=RecordID B=Project C=Status D=Timestamp E=UpdatedBy F=Notes
'   first input field = column 7 (G).  {len(FIELDS)} fields.
'================================================================
Option Explicit
Private Const HDR As Long = 4
Private Const FIRST As Long = 5
Private Const FCOL As Long = 7
Private Const PROJNAME_ADDR As String = "{projname_addr}"

Public Function FieldAddrs() As Variant
    Dim A({len(FIELDS)}) As String
{addr_lines}
    Dim out() As String, i As Long
    ReDim out(1 To {len(FIELDS)})
    For i = 1 To {len(FIELDS)}: out(i) = A(i): Next i
    FieldAddrs = out
End Function

Public Sub Publish_Project()
    Dim db As Worksheet, ck As Worksheet
    Set db = ThisWorkbook.Sheets("Database")
    Set ck = ThisWorkbook.Sheets("Cockpit")
    Dim proj As String: proj = CStr(ck.Range("SelectedProject").Value)
    If Len(Trim(proj)) = 0 Then MsgBox "Select a Project first.", vbExclamation: Exit Sub

    Dim lastRow As Long: lastRow = db.Cells(db.Rows.Count, 1).End(xlUp).Row
    If lastRow < FIRST Then lastRow = HDR

    db.Unprotect          ' Database is locked for hand-editing; unlock for the macro

    ' demote this project's existing Current rows to Historical
    Dim r As Long
    For r = FIRST To lastRow
        If CStr(db.Cells(r, 2).Value) = proj And CStr(db.Cells(r, 3).Value) = "Current" Then
            db.Cells(r, 3).Value = "Historical"
        End If
    Next r

    Dim newRow As Long: newRow = lastRow + 1
    Dim newID As Long
    If lastRow >= FIRST Then
        newID = Application.WorksheetFunction.Max(db.Range("A" & FIRST & ":A" & lastRow)) + 1
    Else
        newID = 1
    End If

    Application.ScreenUpdating = False
    db.Cells(newRow, 1).Value = newID
    db.Cells(newRow, 2).Value = proj
    db.Cells(newRow, 3).Value = "Current"
    db.Cells(newRow, 4).Value = Now
    db.Cells(newRow, 5).Value = Environ("Username")
    db.Cells(newRow, 6).Value = ck.Range("Console_Notes").Value

    Dim addrs As Variant, i As Long
    addrs = FieldAddrs()
    For i = LBound(addrs) To UBound(addrs)
        db.Cells(newRow, FCOL + (i - LBound(addrs))).Value = ck.Range(addrs(i)).Value
    Next i

    EnforceCap db, proj
    db.Protect            ' re-lock the system of record
    Application.ScreenUpdating = True
    Application.Calculate
    MsgBox "Published record #" & newID & " for " & proj & ".", vbInformation
End Sub

Private Sub EnforceCap(db As Worksheet, proj As String)
    Dim cap As Long: cap = CLng(db.Range("MaxRecordsPerProject").Value)
    If cap < 1 Then cap = 1
    Do
        Dim lastRow As Long: lastRow = db.Cells(db.Rows.Count, 1).End(xlUp).Row
        If lastRow < FIRST Then Exit Sub
        Dim cnt As Long
        cnt = Application.WorksheetFunction.CountIf(db.Range("B" & FIRST & ":B" & lastRow), proj)
        If cnt <= cap Then Exit Do
        Dim r As Long, minID As Double, minRow As Long
        minID = 1E+18: minRow = 0
        For r = FIRST To lastRow
            If CStr(db.Cells(r, 2).Value) = proj Then
                If db.Cells(r, 1).Value < minID Then minID = db.Cells(r, 1).Value: minRow = r
            End If
        Next r
        If minRow = 0 Then Exit Do
        ArchiveRow db, minRow
        db.Rows(minRow).Delete
    Loop
End Sub

Private Sub ArchiveRow(db As Worksheet, srcRow As Long)
    Dim ar As Worksheet
    On Error Resume Next: Set ar = ThisWorkbook.Sheets("Archive"): On Error GoTo 0
    If ar Is Nothing Then Exit Sub
    Dim dest As Long: dest = ar.Cells(ar.Rows.Count, 1).End(xlUp).Row + 1
    If dest < 4 Then dest = 4
    ar.Unprotect
    db.Rows(srcRow).Copy
    ar.Rows(dest).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    ar.Protect
End Sub
'''
    m2 = '''Attribute VB_Name = "Module2_Projects"
Option Explicit

Public Sub Add_Project()
    Dim proj As String, phase As String
    proj = Trim(InputBox("New Project name:", "Add Project"))
    If Len(proj) = 0 Then Exit Sub
    phase = Trim(InputBox("Phase No. (optional):", "Add Project", "1"))
    Dim ls As Worksheet: Set ls = ThisWorkbook.Sheets("Lists")
    Dim r As Long: r = ls.Cells(ls.Rows.Count, 1).End(xlUp).Row + 1
    ls.Unprotect
    ls.Cells(r, 1).Value = proj
    ls.Protect
    ' seed a blank Current record
    Dim db As Worksheet: Set db = ThisWorkbook.Sheets("Database")
    Dim lastRow As Long: lastRow = db.Cells(db.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then lastRow = 4
    Dim nr As Long: nr = lastRow + 1
    Dim nid As Long
    If lastRow >= 5 Then nid = Application.WorksheetFunction.Max(db.Range("A5:A" & lastRow)) + 1 Else nid = 1
    db.Unprotect
    db.Cells(nr, 1).Value = nid
    db.Cells(nr, 2).Value = proj
    db.Cells(nr, 3).Value = "Current"
    db.Cells(nr, 4).Value = Now
    db.Cells(nr, 5).Value = Environ("Username")
    db.Cells(nr, 6).Value = "seed (Add Project); Phase " & phase
    db.Protect
    ThisWorkbook.Sheets("Cockpit").Range("SelectedProject").Value = proj
    MsgBox "Added project " & proj & " and selected it.", vbInformation
End Sub

Public Sub Validate_Inputs()
    Dim ck As Worksheet: Set ck = ThisWorkbook.Sheets("Cockpit")
    Dim msg As String
    If Len(Trim(ck.Range("SelectedProject").Value)) = 0 Then msg = msg & "- Project is blank" & vbCrLf
    If ck.Range("Units_Total").Value = 0 Then msg = msg & "- Unit mix has no units" & vbCrLf
    If Len(msg) = 0 Then MsgBox "Required inputs look complete.", vbInformation _
    Else MsgBox "Check these inputs:" & vbCrLf & vbCrLf & msg, vbExclamation
End Sub
'''
    m3 = '''Attribute VB_Name = "Module3_Support"
Option Explicit
Public Const CALENDLY_URL As String = "https://calendly.com/your-org/model-support"
Public Sub Request_Support()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink CALENDLY_URL
    If Err.Number <> 0 Then MsgBox "Open: " & CALENDLY_URL, vbInformation
End Sub
'''
    for fn, ct in [("Module1_Publish.bas", m1), ("Module2_Projects.bas", m2), ("Module3_Support.bas", m3)]:
        with open(os.path.join(vdir, fn), "w", newline="\r\n") as fh:
            fh.write(ct)
    print("Wrote VBA to", vdir)

write_vba()
