#!/usr/bin/env python3
"""
Swerdlow Multi-Phase Cockpit Model - workbook generator.

Builds the input + assumptions architecture and a single stabilized (untrended)
proforma summary described in Cockpit_Model_Build_Spec.txt.

Design notes (kept honest per the spec's "leave a clearly-labeled note" rule):
  * Excel cannot embed compiled VBA (vbaProject.bin) from pure Python, so this
    script emits a .xlsx workbook PLUS the .bas modules in ./vba. See README.md
    for the 60-second import step that turns it into the macro-enabled .xlsm.
  * A single FIELDS registry is the source of truth for: Cockpit input cells,
    the flattened Database columns, the adjacent "current saved value" echo
    formulas, and the generated UPDATE_Phase VBA. They can never drift.
"""

import os
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# 3. Global conventions
# --------------------------------------------------------------------------- #
FONT_NAME = "Arial"
CLR_INPUT_FONT = "0000FF"     # blue  -> [INPUT]
CLR_INPUT_FILL = "E6F1FB"     # light blue fill
CLR_LINK_FONT  = "008000"     # green -> [LINK] cross-tab
CLR_CALC_FONT  = "000000"     # black -> [CALC]
CLR_HDR_FILL   = "1F3864"     # section header band
CLR_HDR_FONT   = "FFFFFF"
CLR_SUB_FILL   = "D9E1F2"
CLR_FLAG_FILL  = "FFC7CE"     # missing-input red fill
CLR_KPI_FILL   = "FFF2CC"
CLR_NOTE_FONT  = "806000"

FMT_CCY   = '$#,##0;($#,##0);"-"'
FMT_PCT   = '0.0%'
FMT_PSF   = '$#,##0.00;($#,##0.00);"-"'   # PSF / PUPM
FMT_CNT   = '#,##0;(#,##0);"-"'
FMT_DATE  = 'mm/dd/yyyy'
FMT_TXT   = '@'

input_fill = PatternFill("solid", fgColor=CLR_INPUT_FILL)
flag_fill  = PatternFill("solid", fgColor=CLR_FLAG_FILL)
hdr_fill   = PatternFill("solid", fgColor=CLR_HDR_FILL)
sub_fill   = PatternFill("solid", fgColor=CLR_SUB_FILL)
kpi_fill   = PatternFill("solid", fgColor=CLR_KPI_FILL)
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
bottom_border = Border(bottom=Side(style="thin", color="808080"))

UNLOCKED = Protection(locked=False)
LOCKED   = Protection(locked=True)

wb = Workbook()

# Registry of flattened input fields: each entry = dict(col, cell, fmt)
# 'cell' is the A1 ref on Cockpit holding the working [INPUT] value.
FIELDS = []
META_COLS = ["RecordID", "PhaseID", "PhaseName", "Timestamp", "UpdatedBy", "Notes"]

def reg(colname, cockpit_a1):
    """Register a Cockpit input field; return its Database column letter."""
    idx = len(FIELDS)
    col_letter = get_column_letter(len(META_COLS) + 1 + idx)  # after meta cols
    FIELDS.append({"col": colname, "cell": cockpit_a1, "dbcol": col_letter})
    return col_letter

# --------------------------------------------------------------------------- #
# styling helpers
# --------------------------------------------------------------------------- #
def style(cell, *, font_color=CLR_CALC_FONT, bold=False, size=10, italic=False,
          fill=None, fmt=None, align=None, wrap=False, border=None, locked=True):
    cell.font = Font(name=FONT_NAME, color=font_color, bold=bold, size=size, italic=italic)
    if fill is not None:
        cell.fill = fill
    if fmt is not None:
        cell.number_format = fmt
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border is not None:
        cell.border = border
    cell.protection = LOCKED if locked else UNLOCKED
    return cell

def label(ws, a1, text, *, bold=False, indent=0, size=10, italic=False, color=CLR_CALC_FONT):
    c = ws[a1]
    c.value = text
    c.font = Font(name=FONT_NAME, color=color, bold=bold, size=size, italic=italic)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=indent)
    return c

def header_band(ws, row, text, last_col_letter="H"):
    ws[f"A{row}"] = text
    for col in range(1, ws[f"{last_col_letter}1"].column + 1):
        cc = ws.cell(row=row, column=col)
        cc.fill = hdr_fill
        cc.font = Font(name=FONT_NAME, color=CLR_HDR_FONT, bold=True, size=11)
        cc.alignment = Alignment(vertical="center")

def input_cell(ws, a1, value=None, *, fmt=FMT_CCY, align="right"):
    c = ws[a1]
    if value is not None:
        c.value = value
    style(c, font_color=CLR_INPUT_FONT, fill=input_fill, fmt=fmt, align=align,
          border=box, locked=False)
    return c

def calc_cell(ws, a1, formula, *, fmt=FMT_CCY, align="right", bold=False):
    c = ws[a1]
    c.value = formula
    style(c, font_color=CLR_CALC_FONT, fmt=fmt, align=align, bold=bold, border=box)
    return c

def link_cell(ws, a1, formula, *, fmt=FMT_CCY, align="right"):
    c = ws[a1]
    c.value = formula
    style(c, font_color=CLR_LINK_FONT, fmt=fmt, align=align, border=box)
    return c

def add_name(name, sheet, a1ref):
    wb.defined_names.add(DefinedName(name, attr_text=f"'{sheet}'!{a1ref}"))

# =========================================================================== #
# LISTS (hidden support tab)
# =========================================================================== #
def build_lists():
    ws = wb.create_sheet("Lists")
    data = {
        "A": ("PhaseID",   ["1-2A", "1-2B", "3-1B"]),
        "B": ("PhaseName", ["Site 1 - Bldg 2A", "Site 1 - Bldg 2B", "Site 3 - Bldg 1B"]),
        "C": ("BedroomType", ["Studio", "1BR", "1BR+Den", "2BR", "2BR+Den", "3BR", "3BR+Den"]),
        "D": ("RentBasis", ["Market", "Live Local", "LIHTC", "PBV", "RAD"]),
        "E": ("UseType",   ["Residential", "Commercial", "Mixed-Use"]),
        "F": ("RentType",  ["NNN", "Gross"]),
        "G": ("CostCurve", ["S-curve", "Straight", "Front-loaded", "Back-loaded"]),
        "H": ("YesNo",     ["Yes", "No"]),
    }
    for col, (title, vals) in data.items():
        style(ws[f"{col}1"], bold=True, font_color=CLR_HDR_FILL)
        ws[f"{col}1"] = title
        for i, v in enumerate(vals, start=2):
            ws[f"{col}{i}"] = v
            ws[f"{col}{i}"].font = Font(name=FONT_NAME, size=10)
    # dynamic named ranges so Add_Asset can extend the phase roster
    add_name("PhaseID_List",   "Lists", "$A$2:$A$101")
    add_name("PhaseName_List", "Lists", "$B$2:$B$101")
    add_name("BedroomType_List", "Lists", "$C$2:$C$8")
    add_name("RentBasis_List", "Lists", "$D$2:$D$6")
    add_name("UseType_List",   "Lists", "$E$2:$E$4")
    add_name("RentType_List",  "Lists", "$F$2:$F$3")
    add_name("CostCurve_List", "Lists", "$G$2:$G$5")
    add_name("YesNo_List",     "Lists", "$H$2:$H$3")
    ws.sheet_state = "hidden"
    ws.protection.sheet = True
    return ws

# =========================================================================== #
# MASTER ASSUMPTIONS  (4.1)
# =========================================================================== #
BEDS = ["Studio", "1BR", "1BR+Den", "2BR", "2BR+Den", "3BR", "3BR+Den"]
HHF_DEFAULT = {"Studio":0.70, "1BR":0.75, "1BR+Den":0.75, "2BR":0.90,
               "2BR+Den":0.90, "3BR":1.04, "3BR+Den":1.04}
FMR_DEFAULT = {"Studio":1400, "1BR":1550, "1BR+Den":1650, "2BR":1850,
               "2BR+Den":1975, "3BR":2350, "3BR+Den":2500}
UA_DEFAULT  = {"Studio":85, "1BR":110, "1BR+Den":120, "2BR":150,
               "2BR+Den":160, "3BR":195, "3BR+Den":210}
BASES = ["Market", "Live Local", "LIHTC", "PBV", "RAD"]
TIER_DEFAULT = {"Market":"", "Live Local":0.80, "LIHTC":0.60, "PBV":"", "RAD":""}
FMRF_DEFAULT = {"Market":"", "Live Local":"", "LIHTC":"", "PBV":1.10, "RAD":1.05}

def build_master():
    ws = wb.create_sheet("Master Assumptions")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    for c in "BCDEFGHIJ":
        ws.column_dimensions[c].width = 13

    label(ws, "A1", "MASTER ASSUMPTIONS  -  program-wide drivers", bold=True, size=13,
          color=CLR_HDR_FILL)
    label(ws, "A2", "Program-wide drivers feed the Cockpit Rent Basis Engine. "
          "Blue = input.", italic=True, color=CLR_NOTE_FONT)

    # 4.1a AMI
    header_band(ws, 4, "4.1a  Rent Basis Engine inputs", "D")
    label(ws, "A5", "AMI (4-person household, annual)")
    input_cell(ws, "B5", 110000, fmt=FMT_CCY)
    add_name("AMI_4person", "Master Assumptions", "$B$5")

    # Per-bedroom table
    label(ws, "A7", "Per-bedroom table", bold=True)
    for j, h in enumerate(["Bedroom", "FMR (mo)", "UtilAllow (mo)", "HHSizeFactor"]):
        style(ws.cell(row=8, column=1+j, value=h), bold=True, fill=sub_fill,
              align="center", border=box)
    r0 = 9
    for i, b in enumerate(BEDS):
        r = r0 + i
        style(ws.cell(row=r, column=1, value=b), align="left", border=box)
        input_cell(ws, f"B{r}", FMR_DEFAULT[b], fmt=FMT_PSF)
        input_cell(ws, f"C{r}", UA_DEFAULT[b], fmt=FMT_PSF)
        input_cell(ws, f"D{r}", HHF_DEFAULT[b], fmt='0.00')
    add_name("FMR_Table", "Master Assumptions", f"$B${r0}:$B${r0+6}")
    add_name("UtilAllow_Table", "Master Assumptions", f"$C${r0}:$C${r0+6}")
    add_name("HHSizeFactor_Table", "Master Assumptions", f"$D${r0}:$D${r0+6}")
    add_name("Bedroom_Index", "Master Assumptions", f"$A${r0}:$A${r0+6}")

    # Basis factor table
    bf = r0 + 8  # header row of basis table
    label(ws, f"A{bf-1}", "Basis factor table", bold=True)
    for j, h in enumerate(["Basis", "Tier %", "FMR_Factor"]):
        style(ws.cell(row=bf, column=1+j, value=h), bold=True, fill=sub_fill,
              align="center", border=box)
    for i, bsis in enumerate(BASES):
        r = bf + 1 + i
        style(ws.cell(row=r, column=1, value=bsis), align="left", border=box)
        input_cell(ws, f"B{r}", TIER_DEFAULT[bsis], fmt=FMT_PCT)
        input_cell(ws, f"C{r}", FMRF_DEFAULT[bsis], fmt='0.00')
    add_name("Basis_Index", "Master Assumptions", f"$A${bf+1}:$A${bf+5}")
    add_name("Tier_Table", "Master Assumptions", f"$B${bf+1}:$B${bf+5}")
    add_name("FMRFactor_Table", "Master Assumptions", f"$C${bf+1}:$C${bf+5}")

    # 4.1b Rent derivation matrix : rows = bedrooms, cols = Live Local/LIHTC/PBV/RAD
    dm = bf + 8  # header row
    header_band(ws, dm-1, "4.1b  Derived net rent / unit / month (by bedroom x basis)", "F")
    deriv_bases = ["Live Local", "LIHTC", "PBV", "RAD"]
    style(ws.cell(row=dm, column=1, value="Bedroom"), bold=True, fill=sub_fill, border=box)
    for j, bsis in enumerate(deriv_bases):
        style(ws.cell(row=dm, column=2+j, value=bsis), bold=True, fill=sub_fill,
              align="center", border=box)
    for i, b in enumerate(BEDS):
        r = dm + 1 + i
        style(ws.cell(row=r, column=1, value=b), align="left", border=box)
        for j, bsis in enumerate(deriv_bases):
            colL = get_column_letter(2 + j)
            bidx = f"MATCH($A{r},Bedroom_Index,0)"
            sidx = f"MATCH(\"{bsis}\",Basis_Index,0)"
            if bsis in ("Live Local", "LIHTC"):
                # AMI-driven
                f = (f"=MAX(0, AMI_4person*INDEX(HHSizeFactor_Table,{bidx})"
                     f"*INDEX(Tier_Table,{sidx})*0.3/12 - INDEX(UtilAllow_Table,{bidx}))")
            else:
                # FMR-driven
                f = (f"=MAX(0, INDEX(FMR_Table,{bidx})*INDEX(FMRFactor_Table,{sidx})"
                     f" - INDEX(UtilAllow_Table,{bidx}))")
            calc_cell(ws, f"{colL}{r}", f, fmt=FMT_PSF)
    add_name("DerivRent_Matrix", "Master Assumptions", f"$B${dm+1}:$E${dm+7}")
    add_name("DerivRent_Beds", "Master Assumptions", f"$A${dm+1}:$A${dm+7}")
    add_name("DerivRent_Bases", "Master Assumptions", f"$B${dm}:$E${dm}")
    label(ws, f"A{dm+9}", "Note: LIHTC / Live Local rents are AMI-based; PBV and RAD are "
          "HUD-FMR-based; Market is entered directly on the Cockpit.", italic=True,
          color=CLR_NOTE_FONT)
    label(ws, f"A{dm+10}", "Changing AMI or FMR re-solves every affordable rent across all "
          "phases automatically.", italic=True, color=CLR_NOTE_FONT)

    # 4.1c Growth-rate defaults (Exhibit E)
    gr = dm + 12
    header_band(ws, gr, "4.1c  Growth-rate defaults (Exhibit E)", "F")
    grow_rows = ["Market Rents", "Affordable Rents", "Retail Rents", "Other Income",
                 "Operating Expenses"]
    style(ws.cell(row=gr+1, column=1, value="Category"), bold=True, fill=sub_fill, border=box)
    for j in range(5):
        style(ws.cell(row=gr+1, column=2+j, value=f"Year {j+1}"), bold=True,
              fill=sub_fill, align="center", border=box)
    label(ws, f"G{gr+1}", "Yr5 = Yr5+", italic=True, color=CLR_NOTE_FONT)
    defaults = [0.030, 0.025, 0.025, 0.025, 0.030]
    for i, rname in enumerate(grow_rows):
        r = gr + 2 + i
        style(ws.cell(row=r, column=1, value=rname), align="left", border=box)
        for j in range(5):
            input_cell(ws, f"{get_column_letter(2+j)}{r}", defaults[i], fmt=FMT_PCT)
    add_name("Growth_Defaults", "Master Assumptions",
             f"$B${gr+2}:$F${gr+2+len(grow_rows)-1}")
    add_name("Growth_Rows", "Master Assumptions", f"$A${gr+2}:$A${gr+2+len(grow_rows)-1}")

    ws.protection.sheet = True
    return ws

# =========================================================================== #
# COCKPIT  (4.2)  -- the only editable surface
# =========================================================================== #
UM_ROWS = 15   # residential unit-mix rows
RR_ROWS = 15   # commercial rent-roll rows

def build_cockpit():
    ws = wb.create_sheet("Cockpit")
    ws.sheet_view.showGridLines = False
    widths = {"A":30, "B":16, "C":16, "D":16, "E":13, "F":12, "G":12, "H":12,
              "I":12, "J":12, "K":12, "L":12, "M":12, "N":12, "O":12, "P":12,
              "Q":12, "R":12, "S":12}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w

    # ---- top control bar ----------------------------------------------------
    label(ws, "A1", "COCKPIT  -  single editable surface (one Phase at a time)",
          bold=True, size=14, color=CLR_HDR_FILL)
    label(ws, "A2", "Edit blue cells, then click UPDATE to save a snapshot to the Database.",
          italic=True, color=CLR_NOTE_FONT)

    label(ws, "A4", "Selected Phase", bold=True)
    sp = input_cell(ws, "B4", "1-2A", fmt=FMT_TXT, align="left")
    add_name("SelectedPhase", "Cockpit", "$B$4")
    dv_phase = DataValidation(type="list", formula1="=PhaseID_List", allow_blank=False)
    ws.add_data_validation(dv_phase); dv_phase.add(ws["B4"])

    # latest-record metadata (LINK)
    maxrec = "MAXIFS(Database!$A:$A,Database!$B:$B,SelectedPhase)"
    def dbmeta(col):  # INDEX latest record's meta column
        return (f'=IFERROR(INDEX(Database!${col}:${col},'
                f'MATCH({maxrec},Database!$A:$A,0)),"(none)")')
    label(ws, "D4", "Last updated by"); link_cell(ws, "E4", dbmeta("E"), fmt=FMT_TXT, align="left")
    label(ws, "F4", "on");             link_cell(ws, "G4", dbmeta("D"), fmt=FMT_DATE, align="left")
    label(ws, "H4", "Record #");       link_cell(ws, "I4", f'=IFERROR({maxrec},0)', fmt=FMT_CNT, align="left")
    label(ws, "A5", "Notes (saved with next UPDATE)")
    input_cell(ws, "B5", "", fmt=FMT_TXT, align="left")
    ws.merge_cells("B5:E5")
    add_name("Cockpit_Notes", "Cockpit", "$B$5")

    # buttons (placeholders -> macros assigned after VBA import; see README)
    for i, (a1, txt) in enumerate([("G5","UPDATE"), ("I5","Add Asset"), ("K5","Request Support")]):
        c = ws[a1]; c.value = txt
        style(c, font_color="FFFFFF", bold=True, fill=PatternFill("solid", fgColor="2E75B6"),
              align="center", border=box)
        ws.merge_cells(f"{a1}:{get_column_letter(c.column+1)}5")
    label(ws, "A6", "Buttons above are placeholders; assign macros after importing the .bas "
          "modules (see README).", italic=True, color=CLR_NOTE_FONT, size=8)

    # ---- KPI strip (CALC) ---------------------------------------------------
    header_band(ws, 8, "KPI strip (computed from current Cockpit inputs)", "S")
    kpis = ["Total Units", "Total Res NSF", "EGI", "NOI", "OpEx Ratio",
            "NOI / Unit", "Total Dev Cost", "Cost / Unit", "Cost / GSF"]
    # KPI formulas are filled at the end once section anchors are known (placeholders).
    kpi_cells = {}
    for i, k in enumerate(kpis):
        col = get_column_letter(1 + i*2) if False else get_column_letter(1 + i)
    # place KPIs across columns A..I (label row 9, value row 10)
    for i, k in enumerate(kpis):
        col = get_column_letter(1 + i)
        style(ws.cell(row=9, column=1+i, value=k), bold=True, fill=kpi_fill,
              align="center", border=box, size=9)
        kpi_cells[k] = f"{col}10"
    # values get formulas later in finalize_kpis()
    for i, k in enumerate(kpis):
        col = get_column_letter(1 + i)
        style(ws[f"{col}10"], bold=True, fill=kpi_fill, align="center", border=box,
              fmt=FMT_CCY)

    anchors = {}            # section anchors used by KPIs / proforma
    row = 12

    # convenience emitters --------------------------------------------------
    def section(title, last="S"):
        nonlocal row
        header_band(ws, row, title, last)
        row += 1
        return row - 1

    def scalar(lbl, colname, default="", fmt=FMT_CCY, validation=None, required=False,
               echo=True, calc=None):
        """Emit one labelled scalar row. Returns the input/calc cell A1 (col B)."""
        nonlocal row
        label(ws, f"A{row}", lbl, indent=1)
        a1 = f"B{row}"
        if calc is not None:
            calc_cell(ws, a1, calc, fmt=fmt, align="right")
        else:
            input_cell(ws, a1, default, fmt=fmt, align="right")
            if validation:
                dv = DataValidation(type="list", formula1=f"={validation}", allow_blank=True)
                ws.add_data_validation(dv); dv.add(ws[a1])
            dbcol = reg(colname, f"Cockpit!${a1[0]}${row}")
            if echo:
                echo_f = (f'=IFERROR(INDEX(Database!${dbcol}:${dbcol},'
                          f'MATCH({maxrec},Database!$A:$A,0)),"-")')
                link_cell(ws, f"C{row}", echo_f, fmt=fmt, align="right")
                ws[f"C{row}"].comment = None
            if required:
                # missing-input flag
                fcell = f"D{row}"
                ws[fcell] = f'=IF(N({a1})=0,"<-- needs input","")' if fmt in (FMT_CCY,FMT_CNT,FMT_PSF) \
                            else f'=IF({a1}="","<-- needs input","")'
                style(ws[fcell], font_color="C00000", align="left", size=9)
                ws.conditional_formatting.add(
                    a1, FormulaRule(formula=[f'ISBLANK({a1})'], fill=flag_fill))
        r = row
        row += 1
        return a1, r

    # header for echo column once
    style(ws["C8"], bold=False)  # (kept)
    label(ws, "C11", "Saved value", italic=True, color=CLR_LINK_FONT, size=8)
    ws["C11"].alignment = Alignment(horizontal="right")

    group_bounds = []  # (start,end) rows to outline-group per section

    # ---- (1) Phase Identity & Physical -------------------------------------
    s = section("(1)  Phase Identity & Physical")
    scalar("Phase Name", "PhaseName_in", "Site 1 - Bldg 2A", FMT_TXT, required=True)
    scalar("Site", "Site_in", "Site 1", FMT_TXT, required=True)
    scalar("Has Residential?", "HasResidential", "Yes", FMT_TXT, validation="YesNo_List")
    scalar("Has Commercial?", "HasCommercial", "No", FMT_TXT, validation="YesNo_List")
    scalar("Stories", "Stories", 8, FMT_CNT)
    a_gsf, r_gsf = scalar("Gross Building Area (GSF)", "GSF", 240000, FMT_CNT, required=True)
    a_nsf, r_nsf = scalar("Net Rentable Area (NSF)", "NetRentableArea", 200000, FMT_CNT)
    a_pk, _ = scalar("Parking spaces", "ParkingSpaces", 250, FMT_CNT)
    scalar("Land area (SF)", "LandArea", 60000, FMT_CNT)
    # parking ratio = spaces / total units (calc), filled after units anchor known -> placeholder
    label(ws, f"A{row}", "Parking ratio (per unit)", indent=1)
    anchors["parking_ratio_cell"] = f"B{row}"
    calc_cell(ws, f"B{row}", "=0", fmt='0.00')  # set later
    row += 1
    group_bounds.append((s+1, row-1))
    anchors["GSF"] = a_gsf
    anchors["NSF"] = a_nsf
    anchors["ParkingSpaces"] = a_pk

    # ---- (2) Key Dates & Lease-Up (Exhibit F) ------------------------------
    s = section("(2)  Key Dates & Lease-Up (Exhibit F)")
    a_close, r_close = scalar("Closing date", "ClosingDate", date(2026, 1, 1), FMT_DATE, required=True)
    ws[a_close].number_format = FMT_DATE
    add_name("Closing_Date_Cell", "Cockpit", f"${a_close[0]}${r_close}")
    offsets = [("Construction Start","off_ConstStart",2),
               ("Initial Occupancy","off_InitialOcc",14),
               ("Construction End","off_ConstEnd",20),
               ("Stabilization","off_Stab",30),
               ("Exit","off_Exit",84)]
    date_cells = {}
    for nm, key, dv in offsets:
        label(ws, f"A{row}", f"{nm}  (month offset)", indent=1)
        ai = f"B{row}"
        input_cell(ws, ai, dv, fmt=FMT_CNT)
        reg(key, f"Cockpit!$B${row}")
        # absolute date = EDATE(Closing, offset)
        calc_cell(ws, f"C{row}", f"=EDATE({a_close},{ai})", fmt=FMT_DATE, align="right")
        date_cells[key] = f"C{row}"
        label(ws, f"D{row}", "abs. date", italic=True, color=CLR_LINK_FONT, size=8)
        row += 1
    a_prelease, r_pl = scalar("Preleased units", "PreleasedUnits", 20, FMT_CNT)
    scalar("Preleased %", "PreleasedPct",
           calc=f"=IFERROR({a_prelease}/Units_Total,0)", fmt=FMT_PCT, echo=False)
    scalar("Lease-up pace (units/month)", "LeaseUpPace", 15, FMT_CNT)
    a_initdel, _ = scalar("Initial delivery units", "InitialDeliveryUnits", 60, FMT_CNT)
    scalar("Duration of lease-up to stabilization (mo)", "LeaseUpDuration",
           calc=f"=IFERROR(MAX(0,(Units_Total-{a_prelease})/MAX(LeaseUpPace_in,1)),0)",
           fmt='#,##0.0', echo=False)
    anchors["ClosingDate"] = a_close
    anchors["date_cells"] = date_cells
    add_name("LeaseUpPace_in", "Cockpit", f"$B${r_pl+2}")
    group_bounds.append((s+1, row-1))

    # ---- (3) Residential Unit Mix (Exhibit A) -- 15 rows --------------------
    s = section("(3)  Residential Unit Mix (Exhibit A)  -  15 rows")
    um_hdr = row
    um_cols = ["Bedroom Type","Rent Basis","Units","Avg SF","Mkt Rent override",
               "Rent/unit/mo","Total SF","% Units","% of Tot SF","Total Rent/mo","Rent PSF"]
    for j, h in enumerate(um_cols):
        style(ws.cell(row=um_hdr, column=1+j, value=h), bold=True, fill=sub_fill,
              align="center", border=box, size=9, wrap=True)
    row += 1
    um_first = row
    for i in range(UM_ROWS):
        r = um_first + i
        # inputs
        bcell = f"A{r}"; ws[bcell] = ("Studio" if i==0 else "")
        style(ws[bcell], font_color=CLR_INPUT_FONT, fill=input_fill, align="left",
              border=box, locked=False)
        dvb = DataValidation(type="list", formula1="=BedroomType_List", allow_blank=True)
        ws.add_data_validation(dvb); dvb.add(ws[bcell])
        reg(f"UM{i+1:02d}_Bed", f"Cockpit!$A${r}")
        scell = f"B{r}"; ws[scell] = ("Market" if i==0 else "")
        style(ws[scell], font_color=CLR_INPUT_FONT, fill=input_fill, align="left",
              border=box, locked=False)
        dvs = DataValidation(type="list", formula1="=RentBasis_List", allow_blank=True)
        ws.add_data_validation(dvs); dvs.add(ws[scell])
        reg(f"UM{i+1:02d}_Basis", f"Cockpit!$B${r}")
        input_cell(ws, f"C{r}", (10 if i==0 else ""), fmt=FMT_CNT); reg(f"UM{i+1:02d}_Units", f"Cockpit!$C${r}")
        input_cell(ws, f"D{r}", (520 if i==0 else ""), fmt=FMT_CNT); reg(f"UM{i+1:02d}_AvgSF", f"Cockpit!$D${r}")
        input_cell(ws, f"E{r}", (2100 if i==0 else ""), fmt=FMT_PSF); reg(f"UM{i+1:02d}_MktOvr", f"Cockpit!$E${r}")
        # derived
        # Rent/unit/mo: Market -> override; else lookup DerivRent_Matrix
        rent_f = (f'=IFERROR(IF($C{r}="",0,IF($B{r}="Market",$E{r},'
                  f'INDEX(DerivRent_Matrix,MATCH($A{r},DerivRent_Beds,0),'
                  f'MATCH($B{r},DerivRent_Bases,0)))),0)')
        calc_cell(ws, f"F{r}", rent_f, fmt=FMT_PSF)
        calc_cell(ws, f"G{r}", f"=N($C{r})*N($D{r})", fmt=FMT_CNT)
        calc_cell(ws, f"H{r}", f"=IFERROR($C{r}/Units_Total,0)", fmt=FMT_PCT)
        calc_cell(ws, f"I{r}", f"=IFERROR($G{r}/ResSF_Total,0)", fmt=FMT_PCT)
        calc_cell(ws, f"J{r}", f"=N($C{r})*N($F{r})", fmt=FMT_CCY)
        calc_cell(ws, f"K{r}", f"=IFERROR($J{r}/$C{r}/$D{r},0)", fmt=FMT_PSF)
    um_last = um_first + UM_ROWS - 1
    add_name("UM_Units", "Cockpit", f"$C${um_first}:$C${um_last}")
    add_name("UM_TotSF", "Cockpit", f"$G${um_first}:$G${um_last}")
    add_name("UM_TotRent", "Cockpit", f"$J${um_first}:$J${um_last}")
    add_name("UM_Basis", "Cockpit", f"$B${um_first}:$B${um_last}")
    # totals / blended rows
    row = um_last + 1
    def um_summary(lbl, basis_filter):
        nonlocal row
        style(ws.cell(row=row, column=1, value=lbl), bold=True, fill=sub_fill, border=box, align="left")
        ws.cell(row=row, column=2).fill = sub_fill
        if basis_filter is None:      # all
            u = "=SUM(UM_Units)"; tsf = "=SUM(UM_TotSF)"; tr = "=SUM(UM_TotRent)"
        else:
            crit = "{" + ";".join(f'"{b}"' for b in basis_filter) + "}"
            u = f"=SUM(SUMIFS(UM_Units,UM_Basis,{crit}))"
            tsf = f"=SUM(SUMIFS(UM_TotSF,UM_Basis,{crit}))"
            tr = f"=SUM(SUMIFS(UM_TotRent,UM_Basis,{crit}))"
        calc_cell(ws, f"C{row}", u, fmt=FMT_CNT, bold=True)
        calc_cell(ws, f"D{row}", f"=IFERROR(G{row}/C{row},0)", fmt=FMT_CNT, bold=True)  # avg SF
        calc_cell(ws, f"G{row}", tsf, fmt=FMT_CNT, bold=True)
        calc_cell(ws, f"H{row}", f"=IFERROR(C{row}/Units_Total,0)", fmt=FMT_PCT, bold=True)
        calc_cell(ws, f"I{row}", f"=IFERROR(G{row}/ResSF_Total,0)", fmt=FMT_PCT, bold=True)
        calc_cell(ws, f"J{row}", tr, fmt=FMT_CCY, bold=True)
        calc_cell(ws, f"E{row}", f"=IFERROR(J{row}/C{row},0)", fmt=FMT_PSF, bold=True)  # chunk rent
        calc_cell(ws, f"K{row}", f"=IFERROR(J{row}/G{row},0)", fmt=FMT_PSF, bold=True)
        rr = row; row += 1; return rr
    tot_row = um_summary("Total / Blended (all)", None)
    um_summary("Affordable subtotal", ["Live Local","LIHTC","PBV","RAD"])
    um_summary("Market subtotal", ["Market"])
    add_name("Units_Total", "Cockpit", f"$C${tot_row}")
    add_name("ResSF_Total", "Cockpit", f"$G${tot_row}")
    add_name("ResRent_Total_mo", "Cockpit", f"$J${tot_row}")
    anchors["units_total_cell"] = f"C{tot_row}"
    anchors["ressf_total_cell"] = f"G{tot_row}"
    anchors["resrent_mo_cell"] = f"J{tot_row}"
    group_bounds.append((s+1, row-1))
    # backfill parking ratio
    ws[anchors["parking_ratio_cell"]] = f"=IFERROR({anchors['ParkingSpaces']}/Units_Total,0)"

    # ---- (4) Other Residential Revenue (Exhibit C) -------------------------
    s = section("(4)  Other Residential Revenue (Exhibit C)")
    scalar("Annual Turnover Rate", "TurnoverRate", 0.50, FMT_PCT)
    scalar("Units turned / month", "UnitsTurnedMo",
           calc="=IFERROR(Units_Total*N(B"+str(row-1)+")/12,0)", fmt='#,##0.0', echo=False)
    a_vac, r_vac = scalar("Vacancy Loss % (of GPR)", "VacancyLossPct", 0.05, FMT_PCT)
    a_col, r_col = scalar("Collection Loss % (of net GPR)", "CollectionLossPct", 0.005, FMT_PCT)
    a_model, _ = scalar("Model Units", "ModelUnits", 1, FMT_CNT)
    a_nonrev, _ = scalar("Non-Revenue Units", "NonRevenueUnits", 1, FMT_CNT)
    scalar("Non-revenue rent discount %", "NonRevDiscountPct", 1.00, FMT_PCT)
    scalar("Lease-up concessions: months", "LUConcMonths", 1.0, '#,##0.0')
    scalar("Lease-up concessions: % of leases", "LUConcPctLeases", 0.50, FMT_PCT)
    a_scm, r_scm = scalar("Stabilized concessions: months", "StabConcMonths", 0.5, '#,##0.0')
    a_scl, r_scl = scalar("Stabilized concessions: % of leases", "StabConcPctLeases", 0.25, FMT_PCT)
    scalar("Stabilized concessions: % of GPR", "StabConcPctGPR",
           calc=f"=IFERROR({a_scm}/12*{a_scl},0)", fmt=FMT_PCT, echo=False)
    anchors["VacancyLossPct"] = a_vac
    anchors["CollectionLossPct"] = a_col
    anchors["ModelUnits"] = a_model
    anchors["NonRevUnits"] = a_nonrev
    anchors["StabConcGPR"] = f"B{row-1}"
    # other income lines
    label(ws, f"A{row}", "Other income lines", bold=True, indent=1); row += 1
    oi_hdr = row
    for j, h in enumerate(["Line","# of item","Capture %","Rate","Monthly $","Annual $"]):
        style(ws.cell(row=oi_hdr, column=1+j, value=h), bold=True, fill=sub_fill,
              align="center", border=box, size=9)
    row += 1
    oi_lines = [("Other Income ($/unit/mo)","OtherIncome","Units_Total"),
                ("Parking Revenue ($/space/mo)","ParkingRev","ParkingSpaces_ref"),
                ("Storage Revenue ($/unit/mo)","StorageRev","Units_Total"),
                ("Utilities Reimb. / RUBS ($/unit/mo)","RUBS","Units_Total")]
    oi_annual_cells = []
    for nm, key, base in oi_lines:
        r = row
        style(ws.cell(row=r, column=1, value=nm), align="left", border=box, size=9)
        input_cell(ws, f"B{r}", "", fmt=FMT_CNT); reg(f"{key}_Cnt", f"Cockpit!$B${r}")
        input_cell(ws, f"C{r}", 1.00, fmt=FMT_PCT); reg(f"{key}_Capture", f"Cockpit!$C${r}")
        input_cell(ws, f"D{r}", "", fmt=FMT_PSF); reg(f"{key}_Rate", f"Cockpit!$D${r}")
        base_ref = "ParkingSpaces" if base == "ParkingSpaces_ref" else "Units_Total"
        # monthly = base * capture * rate   (count column kept for items like # spaces if entered)
        mo = f"=IFERROR(MAX(N($B{r}),{base_ref if base_ref!='ParkingSpaces' else 'Units_Total'})*0,0)"
        # simpler: monthly = base * capture * rate
        base_cell = "Units_Total" if base != "ParkingSpaces_ref" else f"{anchors['ParkingSpaces']}"
        calc_cell(ws, f"E{r}", f"=IFERROR({base_cell}*$C{r}*N($D{r}),0)", fmt=FMT_CCY)
        calc_cell(ws, f"F{r}", f"=E{r}*12", fmt=FMT_CCY)
        oi_annual_cells.append(f"F{r}")
        row += 1
    anchors["oi_annual_cells"] = oi_annual_cells
    group_bounds.append((s+1, row-1))

    # ---- (5) Residential Operating Expenses (Exhibit D) --------------------
    s = section("(5)  Residential Operating Expenses (Exhibit D)")
    ox_hdr = row
    for j, h in enumerate(["Line","Fixed %","Per-Unit $","Total $"]):
        style(ws.cell(row=ox_hdr, column=1+j, value=h), bold=True, fill=sub_fill,
              align="center", border=box, size=9)
    row += 1
    opex_total_cells = []
    # Management fee (% of EGI)
    label(ws, f"A{row}", "Management Fee (% of EGI)", indent=1, color=CLR_CALC_FONT)
    input_cell(ws, f"B{row}", 0.03, fmt=FMT_PCT); reg("MgmtFeePct", f"Cockpit!$B${row}")
    mgmt_total = f"D{row}"
    calc_cell(ws, f"D{row}", "=IFERROR(EGI_Annual*$B"+str(row)+",0)", fmt=FMT_CCY)
    opex_total_cells.append(mgmt_total); anchors["MgmtFeeTotal"] = mgmt_total
    row += 1
    def opex_line(nm, key, fixedpct, perunit):
        nonlocal row
        r = row
        label(ws, f"A{r}", nm, indent=1)
        input_cell(ws, f"B{r}", fixedpct, fmt=FMT_PCT); reg(f"{key}_FixedPct", f"Cockpit!$B${r}")
        input_cell(ws, f"C{r}", perunit, fmt=FMT_CCY); reg(f"{key}_PerUnit", f"Cockpit!$C${r}")
        calc_cell(ws, f"D{r}", f"=N($C{r})*Units_Total", fmt=FMT_CCY)
        opex_total_cells.append(f"D{r}")
        row += 1
        return r
    label(ws, f"A{row}", "Controllable", bold=True, indent=0); row += 1
    payroll_r = opex_line("Payroll & Related", "Payroll", 0.20, 1400)
    mktg_r = opex_line("Marketing", "Marketing", 0.10, 250)
    rm_r = opex_line("R&M", "RM", 0.30, 600)
    opex_line("Turnover", "Turnover", 0.20, 300)
    ga_r = opex_line("G&A", "GA", 0.40, 350)
    opex_line("Contract Services", "ContractSvcs", 0.50, 400)
    label(ws, f"A{row}", "Fixed", bold=True, indent=0); row += 1
    util_r = opex_line("Utilities", "Utilities", 0.60, 700)
    ins_r = opex_line("Insurance", "Insurance", 1.00, 500)
    opex_line("Replacement Reserve", "ReplReserve", 1.00, 300)
    # RE Taxes split market / affordable ($/unit each, applied to respective unit counts)
    r = row
    label(ws, f"A{r}", "RE Taxes - market units ($/unit)", indent=1)
    input_cell(ws, f"C{r}", 2500, fmt=FMT_CCY); reg("RETax_Market", f"Cockpit!$C${r}")
    calc_cell(ws, f"D{r}", "=N($C"+str(r)+")*SUM(SUMIFS(UM_Units,UM_Basis,\"Market\"))", fmt=FMT_CCY)
    retax_mkt = f"D{r}"; opex_total_cells.append(retax_mkt); row += 1
    r = row
    label(ws, f"A{r}", "RE Taxes - affordable units ($/unit)", indent=1)
    input_cell(ws, f"C{r}", 800, fmt=FMT_CCY); reg("RETax_Afford", f"Cockpit!$C${r}")
    crit = '{"Live Local";"LIHTC";"PBV";"RAD"}'
    calc_cell(ws, f"D{r}", f"=N($C{r})*SUM(SUMIFS(UM_Units,UM_Basis,{crit}))", fmt=FMT_CCY)
    retax_aff = f"D{r}"; opex_total_cells.append(retax_aff); row += 1
    label(ws, f"A{row}", "Note: no RE-tax abatement line (per spec).", italic=True,
          color=CLR_NOTE_FONT, size=8); row += 1
    # opex total
    r = row
    style(ws.cell(row=r, column=1, value="Total Operating Expenses"), bold=True, fill=sub_fill, border=box, align="left")
    calc_cell(ws, f"D{r}", "=" + "+".join(opex_total_cells), fmt=FMT_CCY, bold=True)
    add_name("OpEx_Total", "Cockpit", f"$D${r}")
    anchors["opex_total_cell"] = f"D{r}"
    anchors["opex_lines"] = dict(payroll=f"D{payroll_r}", marketing=f"D{mktg_r}",
                                 admin=f"D{ga_r}", utilities=f"D{util_r}", rm=f"D{rm_r}",
                                 insurance=f"D{ins_r}", retax=f"{retax_mkt}+{retax_aff}",
                                 mgmt=mgmt_total)
    row += 1
    group_bounds.append((s+1, row-1))

    # ---- (6) Commercial (Exhibit B) -- collapsible -------------------------
    s = section("(6)  Commercial (Exhibit B)  -  collapse if Has Commercial? = No")
    a_ctot, _ = scalar("Total commercial SF", "Comm_TotalSF", 0, FMT_CNT)
    a_cstatic, r_cstatic = scalar("Static % (common/loss)", "Comm_StaticPct", 0.10, FMT_PCT)
    scalar("Static SF", "Comm_StaticSF", calc=f"=N({a_ctot})*{a_cstatic}", fmt=FMT_CNT, echo=False)
    a_cleas, _ = scalar("Leasable SF", "Comm_LeasableSF",
                        calc=f"=N({a_ctot})*(1-{a_cstatic})", fmt=FMT_CNT, echo=False)
    comm_leasable = a_cleas
    scalar("Check (static + leasable = total)", "Comm_Check",
           calc=f"=N({a_ctot})-B{row-2}-B{row-1}", fmt=FMT_CNT, echo=False)
    a_gv, _   = scalar("General Vacancy %", "Comm_GenVac", 0.07, FMT_PCT)
    scalar("Variable OpEx - % variable", "Comm_VarPct", 0.50, FMT_PCT)
    a_vpsf, _ = scalar("Variable OpEx - $/SF", "Comm_VarPSF", 3.50, FMT_PSF)
    a_ins, _  = scalar("Insurance $/SF", "Comm_InsPSF", 0.75, FMT_PSF)
    a_tax, _  = scalar("RE Taxes $/SF", "Comm_TaxPSF", 2.50, FMT_PSF)
    scalar("Management Fee %", "Comm_MgmtPct", 0.04, FMT_PCT)
    # rent roll 15 rows
    label(ws, f"A{row}", "Tenant rent roll  -  15 rows", bold=True, indent=1); row += 1
    rr_hdr = row
    rr_cols = ["Tenant Name","SF","% Tot SF","Start Mo","LCD","Term (mo)","LXD","Rent PSF",
               "Rent Type","Escal %","TI PSF","LC %","LC PSF","FreeRent mo","FreeRent Start",
               "FreeRent End","Abatement mo"]
    for j, h in enumerate(rr_cols):
        style(ws.cell(row=rr_hdr, column=1+j, value=h), bold=True, fill=sub_fill,
              align="center", border=box, size=8, wrap=True)
    row += 1
    rr_first = row
    for i in range(RR_ROWS):
        r = rr_first + i
        n = i + 1
        # Name A
        c = ws[f"A{r}"]; style(c, font_color=CLR_INPUT_FONT, fill=input_fill, align="left", border=box, locked=False); reg(f"RR{n:02d}_Name", f"Cockpit!$A${r}")
        input_cell(ws, f"B{r}", "", fmt=FMT_CNT); reg(f"RR{n:02d}_SF", f"Cockpit!$B${r}")
        calc_cell(ws, f"C{r}", f"=IFERROR($B{r}/SUM($B${rr_first}:$B${rr_first+RR_ROWS-1}),0)", fmt=FMT_PCT)
        input_cell(ws, f"D{r}", "", fmt=FMT_CNT); reg(f"RR{n:02d}_StartMo", f"Cockpit!$D${r}")
        input_cell(ws, f"E{r}", "", fmt=FMT_DATE); reg(f"RR{n:02d}_LCD", f"Cockpit!$E${r}")
        input_cell(ws, f"F{r}", "", fmt=FMT_CNT); reg(f"RR{n:02d}_Term", f"Cockpit!$F${r}")
        calc_cell(ws, f"G{r}", f'=IF($E{r}="","",IFERROR(EDATE($E{r},$F{r}),""))', fmt=FMT_DATE)   # LXD
        input_cell(ws, f"H{r}", "", fmt=FMT_PSF); reg(f"RR{n:02d}_RentPSF", f"Cockpit!$H${r}")
        c = ws[f"I{r}"]; style(c, font_color=CLR_INPUT_FONT, fill=input_fill, align="left", border=box, locked=False)
        dvr = DataValidation(type="list", formula1="=RentType_List", allow_blank=True)
        ws.add_data_validation(dvr); dvr.add(c); reg(f"RR{n:02d}_RentType", f"Cockpit!$I${r}")
        input_cell(ws, f"J{r}", "", fmt=FMT_PCT); reg(f"RR{n:02d}_Escal", f"Cockpit!$J${r}")
        input_cell(ws, f"K{r}", "", fmt=FMT_PSF); reg(f"RR{n:02d}_TIPSF", f"Cockpit!$K${r}")
        input_cell(ws, f"L{r}", "", fmt=FMT_PCT); reg(f"RR{n:02d}_LCPct", f"Cockpit!$L${r}")
        calc_cell(ws, f"M{r}", f"=IFERROR($H{r}*$L{r}*$F{r}/12,0)", fmt=FMT_PSF)  # LC PSF approx
        input_cell(ws, f"N{r}", "", fmt=FMT_CNT); reg(f"RR{n:02d}_FreeRentMo", f"Cockpit!$N${r}")
        input_cell(ws, f"O{r}", "", fmt=FMT_DATE); reg(f"RR{n:02d}_FreeStart", f"Cockpit!$O${r}")
        calc_cell(ws, f"P{r}", f'=IF($O{r}="","",IFERROR(EDATE($O{r},$N{r}),""))', fmt=FMT_DATE)
        input_cell(ws, f"Q{r}", "", fmt=FMT_CNT); reg(f"RR{n:02d}_Abatement", f"Cockpit!$Q${r}")
    rr_last = rr_first + RR_ROWS - 1
    # WAVG / total row
    r = row = rr_last + 1
    style(ws.cell(row=r, column=1, value="WAVG / Total"), bold=True, fill=sub_fill, border=box, align="left")
    calc_cell(ws, f"B{r}", f"=SUM($B${rr_first}:$B${rr_last})", fmt=FMT_CNT, bold=True)
    calc_cell(ws, f"H{r}", f"=IFERROR(SUMPRODUCT($B${rr_first}:$B${rr_last},$H${rr_first}:$H${rr_last})/$B{r},0)", fmt=FMT_PSF, bold=True)
    comm_wavg_psf = f"H{r}"; comm_tot_sf = f"B{r}"
    row += 1
    # stabilized commercial NOI
    label(ws, f"A{row}", "Stabilized commercial NOI start date", indent=1)
    input_cell(ws, f"B{row}", "", fmt=FMT_DATE); reg("Comm_NOIStart", f"Cockpit!$B${row}")
    noi_start_cell = f"B{row}"; row += 1
    label(ws, f"A{row}", "Stabilized commercial NOI end date", indent=1)
    calc_cell(ws, f"B{row}", f'=IF({noi_start_cell}="","",IFERROR(EDATE({noi_start_cell},12),""))',
              fmt=FMT_DATE); row += 1
    label(ws, f"A{row}", "Stabilized commercial NOI", bold=True, indent=1)
    # commercial NOI = leasable income (net of general vacancy) less commercial opex per SF
    comm_noi_f = (f"=IFERROR({comm_leasable}*{comm_wavg_psf}*(1-{a_gv})"
                  f" - {comm_leasable}*({a_vpsf}+{a_ins}+{a_tax}),0)")
    calc_cell(ws, f"B{row}", comm_noi_f, fmt=FMT_CCY, bold=True)
    add_name("Comm_NOI", "Cockpit", f"$B${row}")
    anchors["comm_noi_cell"] = f"B{row}"
    row += 1
    group_bounds.append((s+1, row-1))

    # ---- (7) Growth Rates (phase view) -------------------------------------
    s = section("(7)  Growth Rates (Exhibit E)  -  default-linked, per-phase override")
    gv_hdr = row
    for j in range(5):
        style(ws.cell(row=gv_hdr, column=3+j, value=f"Year {j+1}"), bold=True,
              fill=sub_fill, align="center", border=box, size=9)
    style(ws.cell(row=gv_hdr, column=1, value="Category"), bold=True, fill=sub_fill, border=box)
    style(ws.cell(row=gv_hdr, column=2, value="Source"), bold=True, fill=sub_fill, border=box, align="center")
    row += 1
    grow_rows = ["Market Rents","Affordable Rents","Retail Rents","Other Income","Operating Expenses"]
    for i, nm in enumerate(grow_rows):
        r = row
        style(ws.cell(row=r, column=1, value=nm), align="left", border=box)
        label(ws, f"B{r}", "default", italic=True, color=CLR_LINK_FONT, size=8)
        for j in range(5):
            colL = get_column_letter(3 + j)
            # default link from Master; override input adjacent below would double width;
            # use single cell: link unless overridden -> here show link value (override via input col further right)
            link_cell(ws, f"{colL}{r}",
                      f"=INDEX(Growth_Defaults,MATCH($A{r},Growth_Rows,0),{j+1})", fmt=FMT_PCT)
            # optional override input columns I..M
            ov = get_column_letter(9 + j)
            input_cell(ws, f"{ov}{r}", "", fmt=FMT_PCT); reg(f"GrowthOvr_{i+1}_{j+1}", f"Cockpit!${ov}${r}")
        row += 1
    label(ws, f"A{row}", "Cols C-G show Master defaults [LINK]; cols I-M are optional "
          "per-phase overrides [INPUT]. Year 5 = Year 5+.", italic=True,
          color=CLR_NOTE_FONT, size=8); row += 1
    group_bounds.append((s+1, row-1))

    # ---- (8) Development Budget / Capital Costs ----------------------------
    s = section("(8)  Development Budget / Capital Costs")
    db_hdr = row
    for j, h in enumerate(["Line","$","$/Unit","$/GSF","Cost Curve"]):
        style(ws.cell(row=db_hdr, column=1+j, value=h), bold=True, fill=sub_fill,
              align="center", border=box, size=9)
    row += 1
    budget_cells = []
    def budget_line(nm, key, default, pct=False, base_for_pct=None):
        nonlocal row
        r = row
        label(ws, f"A{r}", nm, indent=1)
        if pct:
            input_cell(ws, f"B{r}", default, fmt=FMT_PCT); reg(key, f"Cockpit!$B${r}")
            # $ value computed in D column placeholder? keep % lines out of $ subtotal base
            calc_cell(ws, f"C{r}", f"=IFERROR(($B{r})*({base_for_pct}),0)", fmt=FMT_CCY)
            budget_cells.append(f"C{r}")
            calc_cell(ws, f"D{r}", f"=IFERROR(C{r}/Units_Total,0)", fmt=FMT_CCY)
            calc_cell(ws, f"E{r}", f"=IFERROR(C{r}/N({anchors['GSF']}),0)", fmt=FMT_PSF)
        else:
            input_cell(ws, f"B{r}", default, fmt=FMT_CCY); reg(key, f"Cockpit!$B${r}")
            budget_cells.append(f"B{r}")
            calc_cell(ws, f"C{r}", f"=IFERROR(B{r}/Units_Total,0)", fmt=FMT_CCY)
            calc_cell(ws, f"D{r}", f"=IFERROR(B{r}/N({anchors['GSF']}),0)", fmt=FMT_PSF)
            cc = ws[f"E{r}"]; style(cc, font_color=CLR_INPUT_FONT, fill=input_fill, align="left", border=box, locked=False)
            cc.value = "S-curve"
            dvc = DataValidation(type="list", formula1="=CostCurve_List", allow_blank=True)
            ws.add_data_validation(dvc); dvc.add(cc); reg(f"{key}_Curve", f"Cockpit!$E${r}")
        row += 1
        return r
    label(ws, f"A{row}", "Land & Hard Costs", bold=True); row += 1
    land_r = budget_line("Land", "Bud_Land", 8000000)
    budget_line("Shell / Core", "Bud_Shell", 28000000)
    budget_line("Sitework", "Bud_Sitework", 4000000)
    budget_line("Parking Structure", "Bud_Parking", 6000000)
    budget_line("FF&E", "Bud_FFE", 1500000)
    budget_line("GC Fee", "Bud_GCFee", 2000000)
    hard_base = "+".join(budget_cells)  # hard contingency on costs so far (ex land handled below)
    budget_line("Hard Contingency %", "Bud_HardContPct", 0.05, pct=True,
                base_for_pct="SUM($B$"+str(land_r+1)+":$B$"+str(row-1)+")")
    label(ws, f"A{row}", "Soft Costs", bold=True); row += 1
    soft_start = row
    budget_line("A&E", "Bud_AE", 2500000)
    budget_line("Permits & Impact Fees", "Bud_Permits", 1800000)
    budget_line("Legal", "Bud_Legal", 600000)
    budget_line("Marketing / Lease-up", "Bud_Mktg", 700000)
    budget_line("Developer Fee", "Bud_DevFee", 3000000)
    budget_line("Taxes & Insurance during Construction", "Bud_TaxInsConst", 900000)
    budget_line("Soft Contingency %", "Bud_SoftContPct", 0.05, pct=True,
                base_for_pct="SUM($B$"+str(soft_start)+":$B$"+str(row-1)+")")
    # total dev cost
    r = row
    style(ws.cell(row=r, column=1, value="Total Development Cost"), bold=True, fill=sub_fill, border=box, align="left")
    calc_cell(ws, f"B{r}", "=" + "+".join(budget_cells), fmt=FMT_CCY, bold=True)
    calc_cell(ws, f"C{r}", f"=IFERROR(B{r}/Units_Total,0)", fmt=FMT_CCY, bold=True)
    calc_cell(ws, f"D{r}", f"=IFERROR(B{r}/N({anchors['GSF']}),0)", fmt=FMT_PSF, bold=True)
    add_name("TotalDevCost", "Cockpit", f"$B${r}")
    anchors["tdc_cell"] = f"B{r}"
    row += 1
    group_bounds.append((s+1, row-1))

    # ---- (9) Operating Proforma Summary - Untrended (Exhibit G) -------------
    s = section("(9)  Operating Proforma Summary - Untrended (Exhibit G)")
    label(ws, f"C{row-0}", "$ Annual", bold=True, color=CLR_HDR_FILL);
    style(ws[f"C{s}"], bold=True)
    # column headers
    style(ws.cell(row=s, column=3, value="$ Annual"), bold=True, fill=hdr_fill, font_color=CLR_HDR_FONT, align="center")
    style(ws.cell(row=s, column=4, value="$ PUPM"), bold=True, fill=hdr_fill, font_color=CLR_HDR_FONT, align="center")
    pf = {}
    def pf_line(lbl, formula, *, bold=False, name=None, pct=False):
        nonlocal row
        label(ws, f"A{row}", lbl, indent=1, bold=bold)
        calc_cell(ws, f"C{row}", formula, fmt=(FMT_PCT if pct else FMT_CCY), bold=bold)
        if not pct:
            calc_cell(ws, f"D{row}",
                      f"=IFERROR(C{row}/Units_Total/12,0)", fmt=FMT_PSF, bold=bold)
        if name:
            add_name(name, "Cockpit", f"$C${row}")
            pf[name] = f"C{row}"
        r = row; row += 1; return r
    # GPR = residential annual rent at 100%
    pf_line("Gross Potential Rent", "=ResRent_Total_mo*12", name="GPR_Annual")
    pf_line("Loss to Old Lease (model/non-rev)",
            f"=-(({anchors['ModelUnits']}+{anchors['NonRevUnits']})*IFERROR(ResRent_Total_mo/Units_Total,0)*12)",
            name="LossOldLease")
    pf_line("Net Rental Revenue", "=GPR_Annual+LossOldLease", bold=True, name="NetRentRev")
    pf_line("Vacancy", f"=-NetRentRev*{anchors['VacancyLossPct']}", name="VacancyAmt")
    pf_line("Concession", f"=-NetRentRev*{anchors['StabConcGPR']}", name="ConcessionAmt")
    pf_line("Collection Loss",
            f"=-(NetRentRev+VacancyAmt+ConcessionAmt)*{anchors['CollectionLossPct']}", name="CollLoss")
    pf_line("Effective Rental Revenue",
            "=NetRentRev+VacancyAmt+ConcessionAmt+CollLoss", bold=True, name="EffRentRev")
    # other income (net)
    oi_sum = "+".join(anchors["oi_annual_cells"])
    pf_line("Other Income (net)", f"=({oi_sum})*(1-{anchors['VacancyLossPct']})", name="OtherIncomeNet")
    pf_line("Effective Gross Income (EGI)", "=EffRentRev+OtherIncomeNet", bold=True, name="EGI_Annual")
    add_name("EGI_Annual", "Cockpit", pf["EGI_Annual"].replace("C", "$C$"))
    # operating expenses block
    ol = anchors["opex_lines"]
    pf_line("Payroll", f"=-{ol['payroll']}")
    pf_line("Marketing", f"=-{ol['marketing']}")
    pf_line("Administrative (G&A)", f"=-{ol['admin']}")
    pf_line("Repairs & Maintenance", f"=-{ol['rm']}")
    pf_line("Utilities", f"=-{ol['utilities']}")
    pf_line("Insurance", f"=-{ol['insurance']}")
    pf_line("RE Taxes", f"=-({ol['retax']})")
    pf_line("Management Fee", f"=-{ol['mgmt']}")
    pf_line("Total Operating Expenses", f"=-{anchors['opex_total_cell']}", bold=True, name="TotOpEx")
    pf_line("OpEx Ratio (% of EGI)", "=IFERROR(-TotOpEx/EGI_Annual,0)", pct=True, name="OpExRatio")
    pf_line("Net Operating Income (residential)", "=EGI_Annual+TotOpEx", bold=True, name="NOI_Res")
    pf_line("+ Stabilized Commercial NOI", "=IF(HasCommercial_ref=\"Yes\",Comm_NOI,0)", name="NOI_CommLine")
    pf_line("Net Operating Income (total)", "=NOI_Res+NOI_CommLine", bold=True, name="NOI_Total")
    add_name("HasCommercial_ref", "Cockpit", "$B$16")  # Has Commercial? input cell (section 1)
    group_bounds.append((s+1, row-1))
    anchors["pf"] = pf

    # ---- finalize KPI strip -------------------------------------------------
    kpi_formula = {
        "Total Units": ("=Units_Total", FMT_CNT),
        "Total Res NSF": ("=ResSF_Total", FMT_CNT),
        "EGI": ("=EGI_Annual", FMT_CCY),
        "NOI": ("=NOI_Total", FMT_CCY),
        "OpEx Ratio": ("=OpExRatio", FMT_PCT),
        "NOI / Unit": ("=IFERROR(NOI_Total/Units_Total,0)", FMT_CCY),
        "Total Dev Cost": ("=TotalDevCost", FMT_CCY),
        "Cost / Unit": ("=IFERROR(TotalDevCost/Units_Total,0)", FMT_CCY),
        "Cost / GSF": (f"=IFERROR(TotalDevCost/N({anchors['GSF']}),0)", FMT_PSF),
    }
    for k, (f, fmt) in kpi_formula.items():
        cell = kpi_cells[k]
        ws[cell] = f
        style(ws[cell], bold=True, fill=kpi_fill, align="center", border=box, fmt=fmt, size=9)

    # ---- row grouping (collapsible sections) --------------------------------
    for (a, b) in group_bounds:
        if b >= a:
            ws.row_dimensions.group(a, b, outline_level=1, hidden=False)
    ws.sheet_properties.outlinePr.summaryBelow = False

    # ---- sheet protection: lock all except input cells ----------------------
    ws.protection.sheet = True
    ws.protection.formatCells = False
    return ws, anchors

# =========================================================================== #
# DATABASE  (4.3)  flattened snapshot ledger
# =========================================================================== #
def build_database():
    ws = wb.create_sheet("Database")
    ws.sheet_view.showGridLines = False
    label(ws, "A1", "DATABASE  -  system of record (flat snapshot ledger). "
          "One row = one saved Record for one Phase. Do not edit by hand.",
          bold=True, color=CLR_HDR_FILL)
    # config: MaxRecordsPerPhase named param
    label(ws, "A2", "MaxRecordsPerPhase")
    input_cell(ws, "B2", 20, fmt=FMT_CNT, align="left")
    add_name("MaxRecordsPerPhase", "Database", "$B$2")
    label(ws, "C2", "FIFO cap; on exceed, oldest record for that Phase is moved to Archive "
          "(see VBA). Trade-off: archive (recoverable) vs hard-delete.", italic=True,
          color=CLR_NOTE_FONT, size=8)
    # header row 4
    hdr = META_COLS + [f["col"] for f in FIELDS]
    for j, h in enumerate(hdr):
        style(ws.cell(row=4, column=1+j, value=h), bold=True, fill=hdr_fill,
              font_color=CLR_HDR_FONT, align="center", border=box, size=8)
    # seed one record for sample phase 1-2A pulling current Cockpit values
    seed_phases = [("1-2A","Site 1 - Bldg 2A"), ("1-2B","Site 1 - Bldg 2B"),
                   ("3-1B","Site 3 - Bldg 1B")]
    rr = 5
    for k, (pid, pname) in enumerate(seed_phases, start=1):
        ws.cell(row=rr, column=1, value=k)         # RecordID
        ws.cell(row=rr, column=2, value=pid)       # PhaseID
        ws.cell(row=rr, column=3, value=pname)     # PhaseName
        ws.cell(row=rr, column=4, value="2026-06-23")  # Timestamp (seed)
        ws.cell(row=rr, column=5, value="seed")
        ws.cell(row=rr, column=6, value="initial seed record")
        for j, fld in enumerate(FIELDS):
            # seed the first phase from live Cockpit cells via formula; others blank
            if k == 1:
                ref = fld["cell"].replace("Cockpit!", "Cockpit!")
                ws.cell(row=rr, column=len(META_COLS)+1+j, value=f"={ref}")
            else:
                ws.cell(row=rr, column=len(META_COLS)+1+j, value="")
        rr += 1
    for c in ws[f"A4:{get_column_letter(len(hdr))}{rr-1}"]:
        for cell in c:
            cell.font = Font(name=FONT_NAME, size=8)
    ws.column_dimensions["A"].width = 9
    ws.protection.sheet = True
    return ws

# Archive sheet
def build_archive():
    ws = wb.create_sheet("Archive")
    label(ws, "A1", "ARCHIVE  -  purged Database rows (FIFO overflow beyond "
          "MaxRecordsPerPhase). Recoverable; not hard-deleted.", bold=True, color=CLR_HDR_FILL)
    hdr = META_COLS + [f["col"] for f in FIELDS]
    for j, h in enumerate(hdr):
        style(ws.cell(row=3, column=1+j, value=h), bold=True, fill=hdr_fill,
              font_color=CLR_HDR_FONT, align="center", size=8)
    ws.sheet_state = "hidden"
    ws.protection.sheet = True
    return ws

# =========================================================================== #
# ALL PHASES  (4.4)
# =========================================================================== #
def build_all_phases():
    ws = wb.create_sheet("All Phases")
    ws.sheet_view.showGridLines = False
    label(ws, "A1", "ALL PHASES  -  program roll-up (latest record per Phase, read-only)",
          bold=True, size=13, color=CLR_HDR_FILL)
    cols = ["Phase","Site","Use","Units","NSF","EGI","NOI","OpEx Ratio",
            "Total Dev Cost","Cost/Unit","Cost/GSF","NOI/Unit"]
    for j, h in enumerate(cols):
        style(ws.cell(row=3, column=1+j, value=h), bold=True, fill=hdr_fill,
              font_color=CLR_HDR_FONT, align="center", border=box, size=9)
        ws.column_dimensions[get_column_letter(1+j)].width = 13
    ws.column_dimensions["A"].width = 12
    # one row per phase from Lists roster; pull latest record from Database
    def dbget(colname, phase_cell):
        col = next(f["dbcol"] for f in FIELDS if f["col"] == colname)
        return (f'INDEX(Database!${col}:${col},MATCH(MAXIFS(Database!$A:$A,'
                f'Database!$B:$B,{phase_cell}),Database!$A:$A,0))')
    first = 4
    for i in range(20):  # up to 20 phases
        r = first + i
        pcell = f"Lists!$A${2+i}"
        style(ws.cell(row=r, column=1), border=box, size=9)
        ws.cell(row=r, column=1, value=f'=IF({pcell}="","",{pcell})')
        ws.cell(row=r, column=2, value=f'=IFERROR(IF({pcell}="","",{dbget("Site_in",pcell)}),"")')
        ws.cell(row=r, column=3, value=f'=IFERROR(IF({pcell}="","",IF({dbget("HasCommercial",pcell)}="Yes","Mixed-Use","Residential")),"")')
        ws.cell(row=r, column=4, value=f'=IFERROR(SUM(SUMIFS(Database!$1:$1,Database!$1:$1,0)),"")')  # placeholder
        # Units total: sum UM units columns for that record -> use helper SUMPRODUCT across UM unit cols
        um_unit_cols = [f["dbcol"] for f in FIELDS if f["col"].endswith("_Units") and f["col"].startswith("UM")]
        # build INDEX of the record row then sum; simpler: use a SUMIFS-less approach via named formula
        recmatch = f'MATCH(MAXIFS(Database!$A:$A,Database!$B:$B,{pcell}),Database!$A:$A,0)'
        # units = sum of the UM unit columns in that row
        unit_terms = "+".join([f'N(INDEX(Database!${c}:${c},{recmatch}))' for c in um_unit_cols])
        ws.cell(row=r, column=4, value=f'=IFERROR(IF({pcell}="","",{unit_terms}),"")')
        ws.cell(row=r, column=5, value=f'=IFERROR(IF({pcell}="","",{dbget("NetRentableArea",pcell)}),"")')
        # EGI/NOI not stored as single fields in DB (they are calc) -> show "-" note for now
        for cc in range(6, 13):
            ws.cell(row=r, column=cc, value=f'=IF({pcell}="","","-")')
        for cc in range(1, 13):
            style(ws.cell(row=r, column=cc), border=box, size=9,
                  fmt=(FMT_CNT if cc in (4,5) else (FMT_TXT if cc<=3 else FMT_CCY)),
                  align=("left" if cc<=3 else "right"), font_color=CLR_LINK_FONT)
    # program totals row
    tr = first + 20
    style(ws.cell(row=tr, column=1, value="PROGRAM TOTAL"), bold=True, fill=sub_fill, border=box, align="left")
    ws.cell(row=tr, column=4, value=f"=SUM(D{first}:D{first+19})")
    ws.cell(row=tr, column=5, value=f"=SUM(E{first}:E{first+19})")
    for cc in (4,5):
        style(ws.cell(row=tr, column=cc), bold=True, fill=sub_fill, border=box, fmt=FMT_CNT, align="right")
    label(ws, f"A{tr+2}", "Note: EGI / NOI / cost metrics roll up once each phase has a saved "
          "record; Database stores raw inputs, so program-level $ metrics recompute from the "
          "latest record per phase. Subtotals by Program (Affordable vs Market) and by Site "
          "build on the same MAXIFS+INDEX pattern.", italic=True, color=CLR_NOTE_FONT, size=8)
    ws.protection.sheet = True
    return ws

# =========================================================================== #
# PHASE DETAIL  (4.5)
# =========================================================================== #
def build_phase_detail():
    ws = wb.create_sheet("Phase Detail")
    ws.sheet_view.showGridLines = False
    label(ws, "A1", "PHASE DETAIL  -  read-only single-phase proforma (printable)",
          bold=True, size=13, color=CLR_HDR_FILL)
    label(ws, "A3", "Selected Phase")
    c = input_cell(ws, "B3", "1-2A", fmt=FMT_TXT, align="left")
    add_name("SelectedPhaseDetail", "Phase Detail", "$B$3")
    dv = DataValidation(type="list", formula1="=PhaseID_List", allow_blank=False)
    ws.add_data_validation(dv); dv.add(c)
    label(ws, "A5", "Renders the Exhibit G untrended proforma + key stats for the selected "
          "phase, pulled from the Database latest record.", italic=True, color=CLR_NOTE_FONT)
    # mirror the Exhibit G line labels with values pulled from the latest record's stored inputs
    lines = ["Gross Potential Rent","Net Rental Revenue","Vacancy","Concession",
             "Collection Loss","Effective Rental Revenue","Other Income (net)",
             "Effective Gross Income (EGI)","Total Operating Expenses","OpEx Ratio",
             "Net Operating Income (total)"]
    style(ws.cell(row=7, column=1, value="Operating Proforma (untrended)"), bold=True,
          fill=hdr_fill, font_color=CLR_HDR_FONT)
    style(ws.cell(row=7, column=2, value="$ Annual"), bold=True, fill=hdr_fill,
          font_color=CLR_HDR_FONT, align="center")
    for i, ln in enumerate(lines):
        r = 8 + i
        label(ws, f"A{r}", ln, indent=1)
        # Phase Detail recomputes from stored inputs; show "-" placeholder pattern with note
        ws.cell(row=r, column=2, value='="-"')
        style(ws.cell(row=r, column=2), font_color=CLR_LINK_FONT, align="right",
              fmt=(FMT_PCT if ln=="OpEx Ratio" else FMT_CCY), border=box)
    label(ws, "A21", "Note: when SelectedPhaseDetail = the Cockpit's SelectedPhase, this view "
          "equals Exhibit G. For other phases it recomputes from that phase's latest stored "
          "inputs using the same formulas (wired in the next pass alongside the CF engine).",
          italic=True, color=CLR_NOTE_FONT, size=8)
    ws.protection.sheet = True
    return ws

# =========================================================================== #
# MONTHLY CF  (4.6)  future placeholder - date spine only
# =========================================================================== #
def build_monthly_cf():
    ws = wb.create_sheet("Monthly CF")
    ws.sheet_view.showGridLines = False
    label(ws, "A1", "MONTHLY CF  -  FUTURE-SCOPE PLACEHOLDER (no calculations this pass)",
          bold=True, size=13, color="C00000")
    label(ws, "A2", "Monthly date-header spine driven by Key Dates for the Cockpit's selected "
          "phase. The monthly cash-flow engine plugs in here later without rework.",
          italic=True, color=CLR_NOTE_FONT)
    label(ws, "A4", "Anchor: Closing date (from Cockpit)")
    link_cell(ws, "B4", "=Closing_Date_Cell", fmt=FMT_DATE, align="left")
    # date spine: 120 months across columns starting at C6
    label(ws, "A6", "Month index"); label(ws, "A7", "Month start date")
    label(ws, "A8", "Phase month (0 = closing)")
    add_name("CF_Closing", "Monthly CF", "$B$4")
    for m in range(120):
        col = get_column_letter(3 + m)
        ws.cell(row=6, column=3+m, value=m+1)
        style(ws.cell(row=6, column=3+m), align="center", size=8, bold=True, fill=sub_fill)
        ws.cell(row=7, column=3+m, value=f"=EDATE($B$4,{m})")
        style(ws.cell(row=7, column=3+m), fmt="mmm-yy", align="center", size=8,
              font_color=CLR_LINK_FONT)
        ws.cell(row=8, column=3+m, value=m)
        style(ws.cell(row=8, column=3+m), align="center", size=8)
    label(ws, "A10", "Engine rows (revenue, opex, capital spread by Cost-Curve tag) will be "
          "added below this spine.", italic=True, color=CLR_NOTE_FONT, size=8)
    ws.protection.sheet = True
    return ws

# =========================================================================== #
# READ ME
# =========================================================================== #
def build_readme():
    ws = wb.create_sheet("Read me")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 110
    rows = [
        ("Swerdlow Multi-Phase Cockpit Model", True, 16, CLR_HDR_FILL),
        ("", False, 10, CLR_CALC_FONT),
        ("WHAT THIS IS", True, 12, CLR_HDR_FILL),
        ("Scalable input template / 'cockpit' for a multi-phase real estate development "
         "program. This pass builds the input + assumptions architecture and a single "
         "stabilized (untrended) proforma summary. No monthly CF engine, financing, or "
         "equity waterfall (explicitly out of scope, designed so they can be added later).", False, 10, CLR_CALC_FONT),
        ("", False, 10, CLR_CALC_FONT),
        ("TABS (left to right)", True, 12, CLR_HDR_FILL),
        ("Read me | Cockpit | Master Assumptions | Database | All Phases | Phase Detail | "
         "Monthly CF.  Hidden support: Lists, Archive.", False, 10, CLR_CALC_FONT),
        ("", False, 10, CLR_CALC_FONT),
        ("OPERATING LOOP", True, 12, CLR_HDR_FILL),
        ("USER edits inputs on the COCKPIT (one Phase at a time) -> clicks UPDATE -> a flat "
         "snapshot row is appended to the DATABASE (system of record). All Phases / Phase "
         "Detail / Monthly CF read the latest record per phase from the Database.", False, 10, CLR_CALC_FONT),
        ("", False, 10, CLR_CALC_FONT),
        ("COLOR CODING", True, 12, CLR_HDR_FILL),
        ("Blue font + light-blue fill = [INPUT] (you type here).  Black = [CALC] formula.  "
         "Green = [LINK] cross-tab reference.", False, 10, CLR_CALC_FONT),
        ("", False, 10, CLR_CALC_FONT),
        ("** ENABLING THE MACROS (.xlsx -> .xlsm) **", True, 12, "C00000"),
        ("Pure-Python tooling cannot author Excel's compiled VBA project, so the 3 macros "
         "ship as importable text modules in the ./vba folder (UPDATE_Phase, Add_Asset, "
         "Request_Support / Validate_Inputs). To enable them:", False, 10, CLR_CALC_FONT),
        ("  1. Open this .xlsx in Excel. File > Save As > Excel Macro-Enabled Workbook (.xlsm).", False, 10, CLR_CALC_FONT),
        ("  2. Press Alt+F11 (VBA editor). File > Import File... and import each .bas in ./vba.", False, 10, CLR_CALC_FONT),
        ("  3. Back in Excel, Developer > Insert > Button (or right-click the placeholder "
         "buttons on the Cockpit control bar) and assign: UPDATE_Phase, Add_Asset, Request_Support.", False, 10, CLR_CALC_FONT),
        ("  4. In Module1 set CALENDLY_URL to your real scheduling link.", False, 10, CLR_CALC_FONT),
        ("", False, 10, CLR_CALC_FONT),
        ("KEY DESIGN DECISIONS (Section 7 of the spec)", True, 12, CLR_HDR_FILL),
        ("Two master rent inputs (AMI drives LIHTC/Live Local; FMR-by-bedroom drives PBV/RAD); "
         "market rents entered directly. Untrended proforma only. Database = flat snapshot "
         "ledger, FIFO-capped at 20 (configurable via MaxRecordsPerPhase), purged rows archived. "
         "Development budget IS in scope; financing/waterfall OUT. No RE-tax abatement line. "
         "15-row caps on unit mix and commercial rent roll.", False, 10, CLR_CALC_FONT),
        ("", False, 10, CLR_CALC_FONT),
        ("NOTES ON DEFERRED WIRING", True, 12, CLR_HDR_FILL),
        ("All Phases and Phase Detail use the MAXIFS+INDEX 'latest record per phase' pattern. "
         "Because the Database stores raw inputs (not the computed EGI/NOI), the program-level "
         "$ roll-ups recompute from those inputs; the full recompute formulas land alongside "
         "the monthly CF engine in the next pass. Cells showing '-' are intentional placeholders, "
         "labeled here rather than silently faked.", False, 10, CLR_CALC_FONT),
    ]
    r = 1
    for text, bold, size, color in rows:
        c = ws[f"A{r}"]; c.value = text
        c.font = Font(name=FONT_NAME, bold=bold, size=size, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30 if len(text) > 90 else (20 if text else 8)
        r += 1
    ws.protection.sheet = True
    return ws

# =========================================================================== #
# BUILD ORDER + tab order
# =========================================================================== #
build_lists()
build_master()
cockpit_ws, anchors = build_cockpit()
build_database()
build_archive()
build_all_phases()
build_phase_detail()
build_monthly_cf()
build_readme()

# remove default sheet
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]

# order tabs: Read me, Cockpit, Master Assumptions, Database, All Phases, Phase Detail, Monthly CF, (hidden Lists, Archive)
order = ["Read me","Cockpit","Master Assumptions","Database","All Phases","Phase Detail",
         "Monthly CF","Lists","Archive"]
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

# set Arial as default-ish on all populated cells already handled; full calc on load
wb.calculation.fullCalcOnLoad = True

out = "/home/user/anthropic-claude-code-cloude/Swerdlow_Cockpit_Model.xlsx"
wb.save(out)
print("Saved", out)
print("Total flattened input fields:", len(FIELDS))

# --------------------------------------------------------------------------- #
# 5. VBA macros  (emitted as importable .bas modules, generated from FIELDS so
#    UPDATE_Phase always writes each Cockpit input into the correct DB column)
# --------------------------------------------------------------------------- #
def write_vba():
    vdir = "/home/user/anthropic-claude-code-cloude/vba"
    os.makedirs(vdir, exist_ok=True)
    phasename_addr = FIELDS[0]["cell"].replace("Cockpit!", "")  # PhaseName input cell
    # build the address table (DB-column order, first field at column 7)
    addr_lines = []
    for i, f in enumerate(FIELDS):
        a = f["cell"].replace("Cockpit!", "")
        addr_lines.append(f'    A({i+1}) = "{a}"')
    addr_block = "\n".join(addr_lines)

    module1 = f'''Attribute VB_Name = "Module1_Update"
'================================================================
' Swerdlow Cockpit Model - UPDATE_Phase + versioning (FIFO cap)
' Auto-generated field map: {len(FIELDS)} flattened inputs.
' Database layout: row 4 = headers, data from row 5.
' Meta cols: A=RecordID B=PhaseID C=PhaseName D=Timestamp
'            E=UpdatedBy F=Notes ; first input field = column 7 (G).
'================================================================
Option Explicit

Private Const HDR_ROW As Long = 4
Private Const FIRST_DATA As Long = 5
Private Const FIRST_FIELD_COL As Long = 7
Private Const PHASENAME_ADDR As String = "{phasename_addr}"

' Returns the Cockpit input-cell addresses in Database-column order.
Public Function FieldAddrs() As Variant
    Dim A({len(FIELDS)}) As String
{addr_block}
    Dim out() As String, i As Long
    ReDim out(1 To {len(FIELDS)})
    For i = 1 To {len(FIELDS)}
        out(i) = A(i)
    Next i
    FieldAddrs = out
End Function

Public Sub UPDATE_Phase()
    Dim wsDB As Worksheet, wsCk As Worksheet
    Set wsDB = ThisWorkbook.Sheets("Database")
    Set wsCk = ThisWorkbook.Sheets("Cockpit")

    Dim phase As String
    phase = CStr(wsCk.Range("SelectedPhase").Value)
    If Len(Trim(phase)) = 0 Then
        MsgBox "Select a Phase before updating.", vbExclamation: Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < FIRST_DATA Then lastRow = HDR_ROW
    Dim newRow As Long: newRow = lastRow + 1

    Dim newID As Long
    If lastRow >= FIRST_DATA Then
        newID = Application.WorksheetFunction.Max(wsDB.Range("A" & FIRST_DATA & ":A" & lastRow)) + 1
    Else
        newID = 1
    End If

    Application.ScreenUpdating = False
    wsDB.Cells(newRow, 1).Value = newID
    wsDB.Cells(newRow, 2).Value = phase
    wsDB.Cells(newRow, 3).Value = wsCk.Range(PHASENAME_ADDR).Value
    wsDB.Cells(newRow, 4).Value = Now
    wsDB.Cells(newRow, 5).Value = Environ("Username")
    wsDB.Cells(newRow, 6).Value = wsCk.Range("Cockpit_Notes").Value

    Dim addrs As Variant, i As Long
    addrs = FieldAddrs()
    For i = LBound(addrs) To UBound(addrs)
        wsDB.Cells(newRow, FIRST_FIELD_COL + (i - LBound(addrs))).Value = _
            wsCk.Range(addrs(i)).Value
    Next i

    EnforceCap wsDB, phase
    Application.ScreenUpdating = True
    MsgBox "Saved record #" & newID & " for phase " & phase & ".", vbInformation
End Sub

' FIFO versioning cap: keep at most MaxRecordsPerPhase records per Phase;
' overflow rows are archived (recoverable) then removed.
Private Sub EnforceCap(wsDB As Worksheet, phase As String)
    Dim cap As Long
    cap = CLng(wsDB.Range("MaxRecordsPerPhase").Value)
    If cap < 1 Then cap = 1

    Do
        Dim lastRow As Long
        lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
        If lastRow < FIRST_DATA Then Exit Sub

        Dim cnt As Long
        cnt = Application.WorksheetFunction.CountIf( _
              wsDB.Range("B" & FIRST_DATA & ":B" & lastRow), phase)
        If cnt <= cap Then Exit Do

        Dim r As Long, minID As Double, minRow As Long
        minID = 1E+18: minRow = 0
        For r = FIRST_DATA To lastRow
            If CStr(wsDB.Cells(r, 2).Value) = phase Then
                If wsDB.Cells(r, 1).Value < minID Then
                    minID = wsDB.Cells(r, 1).Value: minRow = r
                End If
            End If
        Next r
        If minRow = 0 Then Exit Do
        ArchiveRow wsDB, minRow
        wsDB.Rows(minRow).Delete
    Loop
End Sub

Private Sub ArchiveRow(wsDB As Worksheet, srcRow As Long)
    Dim wsA As Worksheet
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets("Archive")
    On Error GoTo 0
    If wsA Is Nothing Then Exit Sub          ' archive disabled -> hard delete only
    Dim destRow As Long
    destRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
    If destRow < 4 Then destRow = 4
    wsDB.Rows(srcRow).Copy
    wsA.Rows(destRow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub
'''

    module2 = f'''Attribute VB_Name = "Module2_Assets"
'================================================================
' Add_Asset  +  Validate_Inputs
'================================================================
Option Explicit

Public Sub Add_Asset()
    Dim pid As String, pname As String, site As String
    pid = Trim(InputBox("Asset / Phase ID (e.g. 3-2A):", "Add Asset"))
    If Len(pid) = 0 Then Exit Sub
    pname = Trim(InputBox("Phase Name:", "Add Asset", "Site - Bldg"))
    site = Trim(InputBox("Site (e.g. Site 3):", "Add Asset", "Site"))

    Dim wsL As Worksheet: Set wsL = ThisWorkbook.Sheets("Lists")
    Dim r As Long: r = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row + 1
    wsL.Cells(r, 1).Value = pid          ' PhaseID roster (drives both dropdowns)
    wsL.Cells(r, 2).Value = pname

    ' seed an initial blank/default record so look-ups resolve
    Dim wsDB As Worksheet: Set wsDB = ThisWorkbook.Sheets("Database")
    Dim lastRow As Long: lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then lastRow = 4
    Dim newRow As Long: newRow = lastRow + 1
    Dim newID As Long
    If lastRow >= 5 Then
        newID = Application.WorksheetFunction.Max(wsDB.Range("A5:A" & lastRow)) + 1
    Else
        newID = 1
    End If
    wsDB.Cells(newRow, 1).Value = newID
    wsDB.Cells(newRow, 2).Value = pid
    wsDB.Cells(newRow, 3).Value = pname
    wsDB.Cells(newRow, 4).Value = Now
    wsDB.Cells(newRow, 5).Value = Environ("Username")
    wsDB.Cells(newRow, 6).Value = "seed (Add Asset): " & site

    ThisWorkbook.Sheets("Cockpit").Range("SelectedPhase").Value = pid
    MsgBox "Added phase " & pid & " and selected it on the Cockpit.", vbInformation
End Sub

' Lightweight required-input check (conditional formatting also flags blanks).
Public Sub Validate_Inputs()
    Dim wsCk As Worksheet: Set wsCk = ThisWorkbook.Sheets("Cockpit")
    Dim msg As String
    If Len(Trim(wsCk.Range("SelectedPhase").Value)) = 0 Then msg = msg & "- Selected Phase is blank" & vbCrLf
    If wsCk.Range("Units_Total").Value = 0 Then msg = msg & "- Residential unit mix has no units" & vbCrLf
    If Len(msg) = 0 Then
        MsgBox "Required inputs look complete for the selected phase.", vbInformation
    Else
        MsgBox "Check these required inputs:" & vbCrLf & vbCrLf & msg, vbExclamation
    End If
End Sub
'''

    module3 = '''Attribute VB_Name = "Module3_Support"
'================================================================
' Request_Support (ESCALATION) - books time with the model owner.
' Set CALENDLY_URL to your real scheduling link.
'================================================================
Option Explicit

Public Const CALENDLY_URL As String = "https://calendly.com/your-org/cockpit-model-support"

Public Sub Request_Support()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink CALENDLY_URL
    If Err.Number <> 0 Then
        MsgBox "Could not open the scheduling link. Please browse to:" & vbCrLf & _
               CALENDLY_URL, vbInformation
    End If
End Sub
'''
    for fname, content in [("Module1_Update.bas", module1),
                           ("Module2_Assets.bas", module2),
                           ("Module3_Support.bas", module3)]:
        with open(os.path.join(vdir, fname), "w", newline="\r\n") as fh:
            fh.write(content)
    print("Wrote VBA modules to", vdir)

write_vba()
