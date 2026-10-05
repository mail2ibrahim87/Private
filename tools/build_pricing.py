"""Delight International – service pricing workbooks.

Builds the master pricing workbook (all services + combined quotation) and single-service files.
Pilot version: framework (Quote Info, Rates, markups, client quote) + Chemical Cleaning.

Usage:  python3 tools/build_pricing.py pricing/
"""
import os
import runpy
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
VOLUME_BUILDER = os.path.join(HERE, "build_cleaning_template.py")

SERVICES = [  # (code, name, built?)
    ("CC", "Chemical Cleaning", True),
    ("DC", "Decontamination", False),
    ("LOF", "Lube Oil Flushing", False),
    ("AB", "Air Blowing", False),
    ("PIG", "Pigging", False),
    ("TC", "Tank Cleaning", False),
    ("N2L", "Nitrogen / Helium Leak Test", False),
    ("N2F", "Nitrogen Flushing / Purging", False),
]

# ---------------------------------------------------------------- styles (Delight blue / green)
FN = "Arial"
BLUE, GREEN, NAVY = "1E73BE", "3AAA35", "1F3864"
f_title = Font(name=FN, size=16, bold=True, color=NAVY)
f_sub = Font(name=FN, size=10, italic=True, color="595959")
f_sect = Font(name=FN, size=11, bold=True, color="FFFFFF")
f_hdr = Font(name=FN, size=10, bold=True, color="FFFFFF")
f_body = Font(name=FN, size=10)
f_bold = Font(name=FN, size=10, bold=True)
f_in = Font(name=FN, size=10, color="0000FF")
f_link = Font(name=FN, size=10, color="008000")
f_grey = Font(name=FN, size=9, color="808080")
f_red = Font(name=FN, size=10, bold=True, color="C00000")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_RES = PatternFill("solid", fgColor="E2EFDA")
FILL_KEY = PatternFill("solid", fgColor="C6E0B4")
FILL_SECT = PatternFill("solid", fgColor=BLUE)
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_GREEN = PatternFill("solid", fgColor=GREEN)
FILL_MISS = PatternFill("solid", fgColor="F4B084")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_LN = Alignment(horizontal="left", vertical="center")
AL_R = Alignment(horizontal="right", vertical="center")
NF_AED = '#,##0.00;-#,##0.00;"-"'
NF_NUM = '#,##0.##;-#,##0.##;"-"'
NF_PCT = '0.0%;-0.0%;"-"'
NF_M3 = '#,##0.000;-#,##0.000;"-"'


def style(c, font=f_body, fill=None, nf=None, al=None, border=True):
    c.font = font
    if fill: c.fill = fill
    if nf: c.number_format = nf
    if al: c.alignment = al
    if border: c.border = BORDER


def q(n):
    return f"'{n}'" if " " in n else n


class Book:
    def __init__(self):
        self.wb = Workbook()
        self.dvs = {}

    def name(self, n, sheet, ref):
        self.wb.defined_names[n] = DefinedName(n, attr_text=f"{q(sheet)}!{ref}")

    def dv(self, ws, formula, cells):
        key = (ws.title, formula)
        if key not in self.dvs:
            d = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True,
                               errorTitle="Invalid entry", error="Pick a value from the drop-down list.")
            ws.add_data_validation(d)
            self.dvs[key] = d
        self.dvs[key].add(cells)


def label(ws, cell, text, font=f_bold):
    style(ws[cell], font, al=AL_LN); ws[cell].value = text


def inp(ws, cell, value=None, nf=None, al=AL_C):
    style(ws[cell], f_in, FILL_IN, nf=nf, al=al)
    if value is not None:
        ws[cell].value = value


def section_bar(ws, row, text, c1="A", c2="L", fill=FILL_SECT):
    ws.merge_cells(f"{c1}{row}:{c2}{row}")
    c = ws[f"{c1}{row}"]; c.value = text; c.font = f_sect; c.fill = fill; c.alignment = AL_LN
    ws.row_dimensions[row].height = 20


# ================================================================ QUOTE INFO
def build_quote_info(B):
    ws = B.wb.active
    ws.title = "Quote Info"
    ws.sheet_view.showGridLines = False
    for c, w in {"A": 2, "B": 40, "C": 22, "D": 16, "E": 16, "F": 16, "G": 50}.items():
        ws.column_dimensions[c].width = w
    ws["B1"] = "DELIGHT INTERNATIONAL – PRICING: QUOTE INFORMATION"; ws["B1"].font = f_title
    ws["B2"] = ("Fill in this sheet and the Rates sheet first. Yellow = input (blue text) · green = calculated. "
                "Markups, contingency, discount and VAT apply to every service in this workbook."); ws["B2"].font = f_sub

    r = 4
    section_bar(ws, r, "QUOTATION DETAILS", "B", "G")
    fields = [("Quotation No.", "QNo"), ("Revision", "QRev"), ("Date", "QDate"), ("Client", "QClient"),
              ("Attention (client contact)", "QAttn"), ("Project / Plant", "QProject"), ("Location / site", "QLoc"),
              ("Client enquiry / RFQ reference", "QRef"), ("Prepared by", "QPrep")]
    for i, (t, n) in enumerate(fields):
        rr = r + 1 + i
        label(ws, f"B{rr}", t)
        inp(ws, f"C{rr}", al=AL_LN, nf="dd-mmm-yyyy" if n == "QDate" else None)
        ws.merge_cells(f"C{rr}:E{rr}")
        B.name(n, ws.title, f"$C${rr}")

    r = 15
    section_bar(ws, r, "CURRENCY  (rates are entered in AED; the client quote is converted)", "B", "G")
    label(ws, f"B{r+1}", "Quote currency")
    inp(ws, f"C{r+1}", "AED"); B.name("QCurr", ws.title, f"$C${r+1}")
    for j, h in enumerate(["Currency", "AED per 1 unit", "", "", "Source"]):
        if h:
            style(ws.cell(row=r + 2, column=2 + j, value=h), f_hdr, FILL_HDR, al=AL_C)
    cur = [("AED", 1, "Base currency"), ("USD", 3.6725, "UAE central bank peg (AED 3.6725 = USD 1)"),
           ("SAR", "=3.6725/3.75", "Via USD pegs (SAR 3.75 = USD 1) – update if required"),
           ("OMR", "=3.6725/0.3845", "Via USD pegs (OMR 0.3845 = USD 1) – update if required")]
    for i, (c_, v, s_) in enumerate(cur):
        rr = r + 3 + i
        label(ws, f"B{rr}", c_)
        inp(ws, f"C{rr}", v, nf="0.0000")
        ws[f"F{rr}"] = s_; ws[f"F{rr}"].font = f_grey
    B.name("CurrU", ws.title, f"$B${r+3}:$B${r+6}")
    B.name("CurrRate", ws.title, f"$C${r+3}:$C${r+6}")
    B.dv(ws, "=CurrU", f"C{r+1}")
    label(ws, f"B{r+7}", "Exchange rate used (AED per 1 quote-currency unit)")
    ws[f"C{r+7}"] = "=INDEX(CurrRate,MATCH(QCurr,CurrU,0))"; style(ws[f"C{r+7}"], f_bold, FILL_RES, nf="0.0000", al=AL_C)
    B.name("QRate", ws.title, f"$C${r+7}")

    r = 24
    section_bar(ws, r, "WORKING PATTERN", "B", "G")
    label(ws, f"B{r+1}", "Shift length (hours)"); inp(ws, f"C{r+1}", 12); B.name("QShiftH", ws.title, f"$C${r+1}")
    label(ws, f"B{r+2}", "Shifts per day (1 = day only, 2 = 24-hour)"); inp(ws, f"C{r+2}", 2); B.name("QShiftsDay", ws.title, f"$C${r+2}")
    ws[f"G{r+1}"] = "Manpower is costed per shift (rate per person per shift on the Rates sheet)"; ws[f"G{r+1}"].font = f_grey

    r = 28
    section_bar(ws, r, "MARKUP CONTROLS  (selling = cost × (1 + markup))", "B", "G")
    label(ws, f"B{r+1}", "Markup mode")
    inp(ws, f"C{r+1}", "Simple"); B.name("QMkMode", ws.title, f"$C${r+1}")
    ws[f"G{r+1}"] = "Simple = 2 markups (in-house / rented).  Detailed = 5 markups."; ws[f"G{r+1}"].font = f_grey
    ws[f"E{r+1}"] = "Simple"; ws[f"F{r+1}"] = "Detailed"
    for c in ("E", "F"):
        ws[f"{c}{r+1}"].font = Font(name=FN, size=8, color="FFFFFF")
    B.name("MkModeU", ws.title, f"$E${r+1}:$F${r+1}")
    B.dv(ws, "=MkModeU", f"C{r+1}")
    mk = [("SIMPLE – In-house crew & equipment", "MkIn"), ("SIMPLE – Rented crew & equipment, bought-in items", "MkRent"),
          ("DETAILED – In-house manpower", "MkInMP"), ("DETAILED – Rented manpower", "MkRentMP"),
          ("DETAILED – In-house equipment", "MkInEq"), ("DETAILED – Rented equipment", "MkRentEq"),
          ("DETAILED – Chemicals, consumables & third-party services", "MkCons")]
    for i, (t, n) in enumerate(mk):
        rr = r + 2 + i
        label(ws, f"B{rr}", t, f_body)
        inp(ws, f"C{rr}", nf=NF_PCT); B.name(n, ws.title, f"$C${rr}")
    ws[f"G{r+2}"] = "Enter your markups (e.g. 25%). Blank = 0%."; ws[f"G{r+2}"].font = f_red
    # effective matrix
    m0 = r + 10
    label(ws, f"B{m0}", "Effective markup applied (calculated)")
    for j, h in enumerate(["In-house", "Rented", "Third-party"]):
        style(ws.cell(row=m0, column=3 + j, value=h), f_hdr, FILL_HDR, al=AL_C)
    B.name("SrcU", ws.title, f"$C${m0}:$E${m0}")
    cats = ["Manpower", "Equipment", "Chemicals, consumables & services"]
    simple = [["MkIn", "MkRent", "MkRent"], ["MkIn", "MkRent", "MkRent"], ["MkIn", "MkRent", "MkRent"]]
    detail = [["MkInMP", "MkRentMP", "MkCons"], ["MkInEq", "MkRentEq", "MkCons"], ["MkCons", "MkCons", "MkCons"]]
    for i, cname in enumerate(cats):
        rr = m0 + 1 + i
        label(ws, f"B{rr}", cname, f_body)
        for j in range(3):
            c = ws.cell(row=rr, column=3 + j, value=f'=IF(QMkMode="Detailed",N({detail[i][j]}),N({simple[i][j]}))')
            style(c, f_body, FILL_RES, nf=NF_PCT, al=AL_C)
    B.name("MkMatrix", ws.title, f"$C${m0+1}:$E${m0+3}")

    r = m0 + 5
    section_bar(ws, r, "CONTINGENCY, DISCOUNT & VAT", "B", "G")
    for i, (t, n, v, note) in enumerate([
            ("Contingency (% of total cost, internal risk allowance)", "QCont", None, "Added to the price, not marked up"),
            ("Commercial discount (% of price)", "QDisc", None, "Applied to the price before VAT"),
            ("VAT (%)", "QVAT", 0.05, "UAE 5% · KSA 15% · Oman 5% · Qatar 0% – set per quotation")]):
        rr = r + 1 + i
        label(ws, f"B{rr}", t, f_body); inp(ws, f"C{rr}", v, nf=NF_PCT); B.name(n, ws.title, f"$C${rr}")
        ws[f"G{rr}"] = note; ws[f"G{rr}"].font = f_grey

    r = r + 5
    section_bar(ws, r, "COMMERCIAL TERMS  (printed on the client quotation)", "B", "G")
    terms = [("Validity", "QTValidity"), ("Payment terms", "QTPayment"), ("Delivery / mobilisation time", "QTMob"),
             ("Exclusions", "QTExcl"), ("Client scope / supplied by client", "QTClient"), ("Other notes", "QTNotes")]
    for i, (t, n) in enumerate(terms):
        rr = r + 1 + i
        label(ws, f"B{rr}", t)
        inp(ws, f"C{rr}", al=AL_L); ws.merge_cells(f"C{rr}:G{rr}"); ws.row_dimensions[rr].height = 30
        B.name(n, ws.title, f"$C${rr}")

    r = r + len(terms) + 2
    section_bar(ws, r, "COMPANY DETAILS  (letterhead on the client quotation)", "B", "G")
    co = [("Company name", "CoName", "Delight Equipment International L.L.C."),
          ("Address", "CoAddr", "P.O. Box 2932, Abu Dhabi, United Arab Emirates"),
          ("Telephone / fax", "CoTel", "Tel: +971 2 5515641  |  Fax: +971 2 5515643"),
          ("E-mail / website", "CoWeb", "info@delightintl.ae  |  www.delightintl.ae"),
          ("Tagline", "CoTag", "MAINTENANCE SIMPLIFIED"),
          ("Certifications", "CoCert", "ISO 9001 · ISO 14001 · ISO 45001")]
    for i, (t, n, v) in enumerate(co):
        rr = r + 1 + i
        label(ws, f"B{rr}", t)
        inp(ws, f"C{rr}", v, al=AL_LN); ws.merge_cells(f"C{rr}:G{rr}")
        B.name(n, ws.title, f"$C${rr}")
    ws[f"B{r+8}"] = "Company details taken from the Delight brochure – edit for KSA / India offices as required."
    ws[f"B{r+8}"].font = f_grey
    return ws


# ================================================================ RATES
def build_rates(B):
    ws = B.wb.create_sheet("Rates")
    ws.sheet_view.showGridLines = False
    for c, w in {"A": 2, "B": 44, "C": 16, "D": 16, "E": 16, "F": 16, "G": 44}.items():
        ws.column_dimensions[c].width = w
    ws["B1"] = "RATES (COST) – all in AED, excluding markup"; ws["B1"].font = f_title
    ws["B2"] = ("Enter your COST rates once; every service sheet looks them up. Add new items in the blank yellow rows. "
                "Orange cells on a costing sheet = a rate is missing here."); ws["B2"].font = f_sub
    r = 4

    def table(title, hdrs, rows, n_blank, names, nfs):
        nonlocal r
        section_bar(ws, r, title, "B", "G")
        for j, h in enumerate(hdrs):
            style(ws.cell(row=r + 1, column=2 + j, value=h), f_hdr, FILL_HDR, al=AL_C)
        ws.row_dimensions[r + 1].height = 28
        first = r + 2
        allrows = rows + [[None] * len(hdrs)] * n_blank
        for i, row in enumerate(allrows):
            rr = first + i
            for j, v in enumerate(row):
                c = ws.cell(row=rr, column=2 + j, value=v)
                if j == 0:
                    style(c, f_in, FILL_IN, al=AL_LN)
                else:
                    style(c, f_in, FILL_IN, nf=nfs[j], al=AL_C)
        last = first + len(allrows) - 1
        for j, n in enumerate(names):
            if n:
                col = chr(ord("B") + j)
                B.name(n, ws.title, f"${col}${first}:${col}${last}")
        r = last + 2

    grades = ["Project Manager", "Engineer", "Supervisor", "Job Performer / Permit Receiver", "Safety Officer",
              "Technician", "Operator", "Pipe Fitter"]
    table("MANPOWER – cost per person per shift", ["Grade", "In-house cost / shift", "Rented cost / shift", "", "", "Remarks"],
          [[g, None, None, None, None, None] for g in grades], 8, ["MPU", "MPIn", "MPRent"], [None, NF_AED, NF_AED, None, None, "@"])
    eq = ["Chemical circulation pump skid", "Chemical mixing / storage tank", "Air-operated diaphragm pump (air pump)",
          "Steam boiler / steam generator", "Heater", "Gamma jet tank cleaning machine", "Waste / effluent tank",
          "Chemical hoses & fittings (set)", "Air compressor", "Gas monitor / lab test kit"]
    table("EQUIPMENT – cost per unit per day", ["Equipment", "In-house cost / day", "Rented cost / day", "", "", "Remarks"],
          [[e, None, None, None, None, None] for e in eq], 20, ["EQU", "EQIn", "EQRent"], [None, NF_AED, NF_AED, None, None, "@"])
    table("CHEMICALS – cost per drum / IBC", ["Chemical", "Drum size (L)", "Cost per drum", "IBC size (L)", "Cost per IBC", "Remarks"],
          [], 20, ["CHU", "CHDrumL", "CHDrumC", "CHIBCL", "CHIBCC"], [None, NF_NUM, NF_AED, NF_NUM, NF_AED, "@"])
    # default pack sizes
    for rr in range(r - 22, r - 2):
        ws[f"C{rr}"] = 200; ws[f"E{rr}"] = 1000
    table("CONSUMABLES & THIRD-PARTY SERVICES – cost per unit", ["Item", "Unit", "Cost per unit", "", "", "Remarks"],
          [["Waste disposal", "m³", None, None, None, None], ["Waste disposal", "load", None, None, None, None],
           ["Lab analysis", "sample", None, None, None, None], ["Scaffolding", "lump sum", None, None, None, None],
           ["Crane", "day", None, None, None, None]],
          15, ["CNU", "CNUnit", "CNCost"], [None, "@", NF_AED, None, None, "@"])
    table("LOGISTICS – mobilisation / demobilisation (cost per trip)", ["Vehicle", "Cost per trip", "", "", "", "Remarks"],
          [["Low-bed trailer", None, None, None, None, None], ["Flatbed trailer", None, None, None, None, None],
           ["Truck with crane (HIAB)", None, None, None, None, None], ["Pickup", None, None, None, None, None],
           ["Crew bus", None, None, None, None, None]],
          5, ["VHU", "VHCost"], [None, NF_AED, None, None, None, "@"])
    section_bar(ws, r, "SITE EXTRAS", "B", "G")
    ex = [("Accommodation & food – per person per day", "XAcc"), ("Crew transport – per vehicle per day", "XTrans"),
          ("Crew travel / ticket – per person per trip", "XTravel")]
    for i, (t, n) in enumerate(ex):
        rr = r + 1 + i
        label(ws, f"B{rr}", t, f_body); inp(ws, f"C{rr}", nf=NF_AED); B.name(n, ws.title, f"$C${rr}")
    r += 5
    ws[f"B{r}"] = "All rates are left blank intentionally – fill in your own costs. No prices have been assumed."
    ws[f"B{r}"].font = f_red
    return ws


# ================================================================ CHEMICAL CLEANING
def build_cc(B):
    code = "CC"
    names = {"set": f"{code} Vol Settings", "vc": f"{code} Volume Calc", "sum": f"{code} Vol Summary", "lk": f"{code} Lookup"}
    V = runpy.run_path(VOLUME_BUILDER, init_globals={"TARGET_WB": B.wb, "SHEET_NAMES": names})
    wb = B.wb
    vs, vsum = wb[names["set"]], wb[names["sum"]]
    SV = q(names["sum"])
    # tie the volume settings to the quote and fix the job type
    for nm_, f in (("ProjClient", "=QClient"), ("ProjPlant", "=QProject"), ("ProjNo", "=QNo")):
        cell = wb.defined_names[nm_].attr_text.split("!")[1].replace("$", "")
        vs[cell] = f; vs[cell].font = f_link; vs[cell].fill = FILL_RES
    jt = wb.defined_names["JobType"].attr_text.split("!")[1].replace("$", "")
    vs[jt] = "Chemical Cleaning"; vs[jt].fill = FILL_RES; vs[jt].font = Font(name=FN, size=12, bold=True)
    vs.data_validations.dataValidation = [d for d in vs.data_validations.dataValidation if jt not in str(d.sqref)]
    vs[f"D{jt[1:]}"] = "Fixed for the Chemical Cleaning pricing"; vs[f"D{jt[1:]}"].font = f_grey
    wb[names["sum"]].sheet_state = "hidden"
    wb[names["lk"]].sheet_state = "hidden"

    # ------------------------------------------------ costing sheet
    ws = wb.create_sheet(f"{code} Costing")
    ws.sheet_view.showGridLines = False
    for c, w in {"A": 4, "B": 40, "C": 13, "D": 12, "E": 26, "F": 12, "G": 12, "H": 14, "I": 15, "J": 9, "K": 15, "L": 42}.items():
        ws.column_dimensions[c].width = w
    ws["B1"] = "CHEMICAL CLEANING – INTERNAL COSTING  (do not send to client)"; ws["B1"].font = f_title
    ws["B2"] = '="Quotation "&QNo&" Rev "&QRev&"   ·   "&QClient&"   ·   "&QProject'; ws["B2"].font = f_sub
    ws["B3"] = ("Fill the yellow cells. Rates come from the Rates sheet; markups from Quote Info. "
                "Orange = rate missing on the Rates sheet."); ws["B3"].font = f_grey

    # scope block rows 5-12
    section_bar(ws, 5, "SCOPE & DURATION")
    scope = [("Total system volume (all trains)", f"={SV}!$I${V['tot']}", "m³", NF_M3),
             ("Total considered volume for chemical cleaning", f"={SV}!$I${V['ctot']}", "m³", NF_M3),
             ("Number of equipment items / lines with volume", None, "nos", "0"),
             ("Total chemical required (all steps)", f"=SUM({','.join(f'{SV}!$I${x}' for x in V['step_rows'])})", "L", NF_NUM)]
    for i, (t, f, u, nf) in enumerate(scope):
        rr = 6 + i
        label(ws, f"B{rr}", t, f_body)
        if f:
            ws[f"C{rr}"] = f
        style(ws[f"C{rr}"], f_link, FILL_RES, nf=nf, al=AL_C)
        ws[f"D{rr}"] = u; ws[f"D{rr}"].font = f_grey
    ws["F6"] = "← from the CC volume sheets"; ws["F6"].font = f_grey
    label(ws, "B10", "Job duration on site (days)")
    inp(ws, "C10", nf=NF_NUM); B.name("CC_Days", ws.title, "$C$10")
    ws["D10"] = "days"; ws["D10"].font = f_grey
    ws["F10"] = "Total days only (rig-up to rig-down). Line items can override days individually."; ws["F10"].font = f_grey
    label(ws, "B11", "Shifts per day (from Quote Info)", f_body)
    ws["C11"] = "=QShiftsDay"; style(ws["C11"], f_link, FILL_RES, al=AL_C)
    label(ws, "B12", "Client price breakdown")
    inp(ws, "C12", "Lump sum"); ws.merge_cells("C12:D12"); B.name("CC_Mode", ws.title, "$C$12")
    ws["F12"] = "Lump sum · Per unit (train) · Per equipment – split by considered-volume share"; ws["F12"].font = f_grey
    ws["N12"] = "Lump sum"; ws["O12"] = "Per unit (train)"; ws["P12"] = "Per equipment"
    for c in "NOP":
        ws[f"{c}12"].font = Font(name=FN, size=8, color="FFFFFF")
    B.name("ModeU", ws.title, "$N$12:$P$12")
    B.dv(ws, "=ModeU", "C12")
    dv_num = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                            showErrorMessage=True, errorTitle="Number", error="Enter a positive number.")
    ws.add_data_validation(dv_num); dv_num.add("C10")

    SUM_TOP = 14   # cost summary block (filled after the sections are laid out)
    row = 32
    subtotals = []   # (label, cost cell, selling cell)
    HEAD = ["#", "Item", "Source", "Qty", "Unit", "Days / qty", "Total units", "Unit cost (AED)", "Cost (AED)", "Markup", "Selling (AED)", "Notes"]

    def head(title, hdr=HEAD):
        nonlocal row
        section_bar(ws, row, title)
        row += 1
        for j, h in enumerate(hdr):
            style(ws.cell(row=row, column=1 + j, value=h), f_hdr, FILL_HDR, al=AL_C)
        ws.row_dimensions[row].height = 26
        row += 1

    def money(r, cat):
        ws[f"I{r}"] = f"=N(G{r})*N(H{r})"
        ws[f"J{r}"] = f'=IF(I{r}=0,0,IFERROR(INDEX(MkMatrix,{cat},MATCH(C{r},SrcU,0)),0))'
        ws[f"K{r}"] = f"=I{r}*(1+J{r})"
        for c, nf in (("I", NF_AED), ("J", NF_PCT), ("K", NF_AED)):
            style(ws[f"{c}{r}"], f_body, FILL_RES, nf=nf, al=AL_C)
        style(ws[f"L{r}"], f_grey, al=AL_L)
        ws.conditional_formatting.add(f"H{r}", FormulaRule(formula=[f'AND(N(G{r})>0,N(H{r})=0)'], fill=FILL_MISS))

    def subtotal(lab, r1, r2):
        nonlocal row
        label(ws, f"B{row}", f"Subtotal – {lab}")
        for c in "IK":
            ws[f"{c}{row}"] = f"=SUM({c}{r1}:{c}{r2})"; style(ws[f"{c}{row}"], f_bold, FILL_KEY, nf=NF_AED, al=AL_C)
        subtotals.append((lab, f"I{row}", f"K{row}"))
        row += 2

    def src_cell(r, default, choices="=SrcU"):
        inp(ws, f"C{r}", default); B.dv(ws, choices, f"C{r}")

    # 1. Manpower
    head("1.  MANPOWER  (persons per shift × days × shifts per day × cost per shift)")
    grades = ["Project Manager", "Supervisor", "Job Performer / Permit Receiver", "Safety Officer", "Technician",
              "Operator", "Pipe Fitter", None, None, None, None, None]
    m1 = row
    for i, g in enumerate(grades):
        r = row
        ws[f"A{r}"] = i + 1; style(ws[f"A{r}"], f_grey, al=AL_C)
        inp(ws, f"B{r}", g, al=AL_LN); B.dv(ws, "=MPU", f"B{r}")
        src_cell(r, "In-house")
        inp(ws, f"D{r}", nf=NF_NUM); dv_num.add(f"D{r}")
        ws[f"E{r}"] = "persons / shift"; style(ws[f"E{r}"], f_grey, al=AL_C)
        inp(ws, f"F{r}", nf=NF_NUM); dv_num.add(f"F{r}")
        ws[f"G{r}"] = f'=IF(N(D{r})=0,"",D{r}*IF(ISNUMBER(F{r}),F{r},N(CC_Days))*N(QShiftsDay))'
        style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
        ws[f"H{r}"] = (f'=IF(B{r}="","",IFERROR(IF(C{r}="Rented",INDEX(MPRent,MATCH(B{r},MPU,0)),'
                       f'INDEX(MPIn,MATCH(B{r},MPU,0))),0))')
        style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
        money(r, 1)
        if i == 0:
            ws[f"L{r}"] = "Days blank = job duration. Total units = person-shifts."
        row += 1
    m2 = row - 1
    B.name("CC_Crew", ws.title, f"$D${m1}:$D${m2}")
    subtotal("Manpower", m1, m2)

    # 2. Equipment
    head("2.  EQUIPMENT  (qty × days × cost per day)")
    eqs = ["Chemical circulation pump skid", "Chemical mixing / storage tank", "Air-operated diaphragm pump (air pump)",
           "Steam boiler / steam generator", "Heater", "Gamma jet tank cleaning machine", "Waste / effluent tank",
           "Chemical hoses & fittings (set)", "Gas monitor / lab test kit", None, None, None, None, None, None]
    e1 = row
    for i, e in enumerate(eqs):
        r = row
        ws[f"A{r}"] = i + 1; style(ws[f"A{r}"], f_grey, al=AL_C)
        inp(ws, f"B{r}", e, al=AL_LN); B.dv(ws, "=EQU", f"B{r}")
        src_cell(r, "In-house")
        inp(ws, f"D{r}", nf=NF_NUM); dv_num.add(f"D{r}")
        ws[f"E{r}"] = "units"; style(ws[f"E{r}"], f_grey, al=AL_C)
        inp(ws, f"F{r}", nf=NF_NUM); dv_num.add(f"F{r}")
        ws[f"G{r}"] = f'=IF(N(D{r})=0,"",D{r}*IF(ISNUMBER(F{r}),F{r},N(CC_Days)))'
        style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
        ws[f"H{r}"] = (f'=IF(B{r}="","",IFERROR(IF(C{r}="Rented",INDEX(EQRent,MATCH(B{r},EQU,0)),'
                       f'INDEX(EQIn,MATCH(B{r},EQU,0))),0))')
        style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
        money(r, 2)
        if i == 0:
            ws[f"L{r}"] = "Days blank = job duration. Total units = unit-days."
        row += 1
    e2 = row - 1
    subtotal("Equipment", e1, e2)

    # 3. Chemicals – one row per chemical step (litres from the volume sheets)
    head("3.  CHEMICALS  (litres per step from the volume sheets → drums / IBCs × cost per pack)",
         ["#", "Step (from CC Vol Settings)", "Source", "Litres", "Chemical (Rates sheet)", "Pack", "No. of packs",
          "Cost / pack (AED)", "Cost (AED)", "Markup", "Selling (AED)", "Notes"])
    c1 = row
    for k, (sr, smr) in enumerate(zip(V["STEP_ROWS"], V["step_rows"])):
        r = row
        ws[f"A{r}"] = k + 1; style(ws[f"A{r}"], f_grey, al=AL_C)
        ws[f"B{r}"] = f'=IF({q(names["set"])}!$B${sr}="","(step not used)",{q(names["set"])}!$B${sr})'
        style(ws[f"B{r}"], f_link, FILL_RES, al=AL_LN)
        src_cell(r, "Third-party")
        ws[f"D{r}"] = f"={SV}!$I${smr}"; style(ws[f"D{r}"], f_link, FILL_RES, nf=NF_NUM, al=AL_C)
        inp(ws, f"E{r}", al=AL_LN); B.dv(ws, "=CHU", f"E{r}")
        inp(ws, f"F{r}", "Drum"); B.dv(ws, '"Drum,IBC"', f"F{r}")
        size = f'IF(F{r}="IBC",INDEX(CHIBCL,MATCH(E{r},CHU,0)),INDEX(CHDrumL,MATCH(E{r},CHU,0)))'
        ws[f"G{r}"] = f'=IF(OR(N(D{r})=0,E{r}=""),"",IFERROR(ROUNDUP(D{r}/{size},0),"check pack size"))'
        style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
        ws[f"H{r}"] = (f'=IF(E{r}="","",IFERROR(IF(F{r}="IBC",INDEX(CHIBCC,MATCH(E{r},CHU,0)),'
                       f'INDEX(CHDrumC,MATCH(E{r},CHU,0))),0))')
        style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
        money(r, 3)
        if k == 0:
            ws[f"L{r}"] = "Pick the chemical & pack; packs rounded up"
        ws.conditional_formatting.add(f"E{r}", FormulaRule(formula=[f'AND(N(D{r})>0,E{r}="")'], fill=FILL_MISS))
        row += 1
    c2 = row - 1
    subtotal("Chemicals", c1, c2)

    # 4. Consumables & third-party services
    head("4.  CONSUMABLES, WASTE DISPOSAL & THIRD-PARTY SERVICES")
    cons = [("Waste disposal", "m³"), ("Lab analysis", "sample"), ("Scaffolding", "lump sum"), ("Crane", "day")] + [(None, None)] * 8
    n1 = row
    for i, (it, un) in enumerate(cons):
        r = row
        ws[f"A{r}"] = i + 1; style(ws[f"A{r}"], f_grey, al=AL_C)
        inp(ws, f"B{r}", it, al=AL_LN); B.dv(ws, "=CNU", f"B{r}")
        src_cell(r, "Third-party")
        inp(ws, f"D{r}", nf=NF_NUM); dv_num.add(f"D{r}")
        inp(ws, f"E{r}", un)
        style(ws[f"F{r}"], f_grey)
        ws[f"G{r}"] = f'=IF(N(D{r})=0,"",D{r})'; style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
        ws[f"H{r}"] = (f'=IF(B{r}="","",IFERROR(INDEX(CNCost,MATCH(1,INDEX((CNU=B{r})*(CNUnit=E{r}),0),0)),'
                       f'IFERROR(INDEX(CNCost,MATCH(B{r},CNU,0)),0)))')
        style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
        money(r, 3)
        row += 1
    ws[f"L{n1}"] = f'="Considered volume = "&TEXT(C7,"#,##0.0")&" m³ × no. of fills"'
    n2 = row - 1
    subtotal("Consumables & services", n1, n2)

    # 5. Mobilisation / demobilisation
    head("5.  MOBILISATION / DEMOBILISATION  (vehicles × trips × cost per trip, crew travel)")
    v1 = row
    vehicles = ["Low-bed trailer", "Flatbed trailer", "Truck with crane (HIAB)", "Pickup", None, None]
    for i, vh in enumerate(vehicles):
        r = row
        ws[f"A{r}"] = i + 1; style(ws[f"A{r}"], f_grey, al=AL_C)
        inp(ws, f"B{r}", vh, al=AL_LN); B.dv(ws, "=VHU", f"B{r}")
        src_cell(r, "Rented")
        inp(ws, f"D{r}", nf=NF_NUM); dv_num.add(f"D{r}")
        ws[f"E{r}"] = "vehicles"; style(ws[f"E{r}"], f_grey, al=AL_C)
        inp(ws, f"F{r}", 2, nf=NF_NUM); dv_num.add(f"F{r}")
        ws[f"G{r}"] = f'=IF(N(D{r})=0,"",D{r}*N(F{r}))'; style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
        ws[f"H{r}"] = f'=IF(B{r}="","",IFERROR(INDEX(VHCost,MATCH(B{r},VHU,0)),0))'
        style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
        money(r, 2)
        if i == 0:
            ws[f"L{r}"] = "Days / qty = no. of trips (2 = mob + demob)"
        row += 1
    r = row
    ws[f"A{r}"] = len(vehicles) + 1; style(ws[f"A{r}"], f_grey, al=AL_C)
    label(ws, f"B{r}", "Crew travel / tickets", f_body)
    src_cell(r, "Third-party")
    inp(ws, f"D{r}", nf=NF_NUM); dv_num.add(f"D{r}")
    ws[f"E{r}"] = "persons"; style(ws[f"E{r}"], f_grey, al=AL_C)
    inp(ws, f"F{r}", 2, nf=NF_NUM)
    ws[f"G{r}"] = f'=IF(N(F{r})=0,"",IF(ISNUMBER(D{r}),D{r},SUM(CC_Crew)*N(QShiftsDay))*F{r})'
    style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
    ws[f"H{r}"] = "=N(XTravel)"; style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
    money(r, 1)
    ws[f"L{r}"] = "Persons blank = total crew (persons per shift × shifts per day); trips = 2"
    row += 1
    v2 = row - 1
    subtotal("Mobilisation / demobilisation", v1, v2)

    # 6. Site extras
    head("6.  SITE EXTRAS")
    x1 = row
    r = row
    ws[f"A{r}"] = 1; style(ws[f"A{r}"], f_grey, al=AL_C)
    label(ws, f"B{r}", "Accommodation & food", f_body)
    src_cell(r, "Third-party")
    inp(ws, f"D{r}", nf=NF_NUM)
    ws[f"E{r}"] = "persons"; style(ws[f"E{r}"], f_grey, al=AL_C)
    inp(ws, f"F{r}", nf=NF_NUM)
    ws[f"G{r}"] = f'=IF(ISNUMBER(D{r}),D{r},SUM(CC_Crew)*N(QShiftsDay))*IF(ISNUMBER(F{r}),F{r},N(CC_Days))'
    style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
    ws[f"H{r}"] = "=N(XAcc)"; style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
    money(r, 1)
    ws[f"L{r}"] = "Persons blank = total crew; days blank = job duration. Enter 0 persons if not required."
    row += 1
    r = row
    ws[f"A{r}"] = 2; style(ws[f"A{r}"], f_grey, al=AL_C)
    label(ws, f"B{r}", "Crew transport (bus / pickup)", f_body)
    src_cell(r, "Rented")
    inp(ws, f"D{r}", nf=NF_NUM)
    ws[f"E{r}"] = "vehicles"; style(ws[f"E{r}"], f_grey, al=AL_C)
    inp(ws, f"F{r}", nf=NF_NUM)
    ws[f"G{r}"] = f'=IF(N(D{r})=0,"",D{r}*IF(ISNUMBER(F{r}),F{r},N(CC_Days)))'
    style(ws[f"G{r}"], f_body, FILL_RES, nf=NF_NUM, al=AL_C)
    ws[f"H{r}"] = "=N(XTrans)"; style(ws[f"H{r}"], f_link, FILL_RES, nf=NF_AED, al=AL_C)
    money(r, 2)
    row += 1
    x2 = row - 1
    subtotal("Site extras", x1, x2)
    LAST = row

    # ------------------------------------------------ cost summary block
    r = SUM_TOP
    section_bar(ws, r, "COST & PRICE SUMMARY (AED)")
    for j, h in enumerate(["", "Section", "", "", "", "", "", "", "Cost", "", "Selling", ""]):
        if h:
            style(ws.cell(row=r + 1, column=1 + j, value=h), f_hdr, FILL_HDR, al=AL_C)
    for i, (lab, cc_, sc) in enumerate(subtotals):
        rr = r + 2 + i
        label(ws, f"B{rr}", lab, f_body)
        ws[f"I{rr}"] = f"={cc_}"; ws[f"K{rr}"] = f"={sc}"
        for c in "IK":
            style(ws[f"{c}{rr}"], f_body, FILL_RES, nf=NF_AED, al=AL_C)
    rr = r + 2 + len(subtotals)
    s1, s2 = r + 2, rr - 1
    lines = [
        ("TOTAL DIRECT COST / SELLING BEFORE CONTINGENCY", f"=SUM(I{s1}:I{s2})", f"=SUM(K{s1}:K{s2})", True),
        ("Contingency (on cost, Quote Info)", None, f"=I{rr}*N(QCont)", False),
        ("Price before discount", None, f"=K{rr}+K{rr+1}", False),
        ("Discount (Quote Info)", None, f"=-K{rr+2}*N(QDisc)", False),
        ("NET PRICE EXCL. VAT (AED)", None, f"=K{rr+2}+K{rr+3}", True),
        ("Gross margin on net price", None, f'=IF(K{rr+4}=0,0,(K{rr+4}-I{rr})/K{rr+4})', False),
        ('="NET PRICE EXCL. VAT ("&QCurr&")"', None, f"=K{rr+4}/QRate", True),
    ]
    for i, (lab, fi, fk, strong) in enumerate(lines):
        x = rr + i
        label(ws, f"B{x}", lab, f_bold if strong else f_body)
        if fi:
            ws[f"I{x}"] = fi
        ws[f"K{x}"] = fk
        for c in "IK":
            style(ws[f"{c}{x}"], f_bold if strong else f_body, FILL_KEY if strong else FILL_RES,
                  nf=NF_PCT if "margin" in lab else NF_AED, al=AL_C)
    B.name("CC_Cost", ws.title, f"$I${rr}")
    B.name("CC_PriceGross", ws.title, f"$K${rr+2}")
    B.name("CC_Discount", ws.title, f"$K${rr+3}")
    B.name("CC_Net", ws.title, f"$K${rr+4}")
    assert rr + len(lines) < 32, "summary block overlaps sections"
    ws.freeze_panes = "C5"
    ws.print_area = f"A1:L{LAST}"
    ws.page_setup.orientation = "landscape"; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0

    # count of items with volume
    alloc = build_alloc(B, code, names, V)
    ws["C8"] = f"=COUNTIF({q(alloc.title)}!$H:$H,\">0\")"
    build_client_quote(B, code, "Chemical Cleaning Services", alloc,
                       "Chemical cleaning of equipment and piping as per scope, including chemicals, manpower, "
                       "equipment, mobilisation and demobilisation.")
    return ws


def build_alloc(B, code, names, V):
    """Hidden sheet: per-equipment / per-train price split by considered-volume share, and the display list."""
    ws = B.wb.create_sheet(f"{code} Alloc")
    VC = q(names["vc"]); SV = q(names["sum"])
    ws["A1"] = "Hidden working sheet – price allocation by considered volume"; ws["A1"].font = f_bold
    hd = ["Kind", "Item", "Train", "Tag", "Train name", "Considered m³", "Flag", "Volume (>0)", "Running no.", "Price (AED)"]
    for j, h in enumerate(hd):
        ws.cell(row=3, column=1 + j, value=h).font = f_bold
    r = 4
    TR = V["TR"]
    for kind, k, tagrow, consrow in V["ITEMS"]:
        for j, col in enumerate(TR):
            ws[f"A{r}"] = kind; ws[f"B{r}"] = k; ws[f"C{r}"] = j + 1
            ws[f"D{r}"] = f'=IF({VC}!{col}{tagrow}="","{kind} {k}",{VC}!{col}{tagrow})'
            ws[f"E{r}"] = f"=TrainName{j+1}"
            ws[f"F{r}"] = f"=N({VC}!{col}{consrow})"
            ws[f"G{r}"] = f"=IF(F{r}>0,1,0)"
            ws[f"H{r}"] = f"=F{r}*G{r}"
            ws[f"I{r}"] = f"=IF(G{r}=1,COUNTIF($G$4:G{r},1),\"\")"
            r += 1
    last = r - 1
    ws["L3"] = "Total considered m³"; ws["L4"] = f"=SUM(F4:F{last})"
    for rr in range(4, last + 1):
        ws[f"J{rr}"] = f"=IF($L$4=0,0,{code}_Net*F{rr}/$L$4)"
    # per-train list
    ws["N3"] = "Train"; ws["O3"] = "Name"; ws["P3"] = "Considered m³"; ws["Q3"] = "Running no."; ws["R3"] = "Price (AED)"
    for j, col in enumerate(TR):
        rr = 4 + j
        ws[f"N{rr}"] = j + 1; ws[f"O{rr}"] = f"=TrainName{j+1}"
        ws[f"P{rr}"] = f"=N({SV}!{col}{V['ctot']})"
        ws[f"Q{rr}"] = f'=IF(P{rr}>0,COUNTIF($P$4:P{rr},">0"),"")'
        ws[f"R{rr}"] = f"=IF($L$4=0,0,{code}_Net*P{rr}/$L$4)"
    B.name(f"{code}_NItems", ws.title, f"$L$7")
    ws["L6"] = "Display rows needed"
    ws["L7"] = (f'=IF({code}_Mode="Per equipment",MAX(0,MAX(I4:I{last})),IF({code}_Mode="Per unit (train)",'
                f'MAX(0,MAX(Q4:Q8)),1))')
    # display list for the client quote
    ws["T3"] = "Disp #"; ws["U3"] = "Description"; ws["V3"] = "Unit / train"; ws["W3"] = "Considered m³"; ws["X3"] = "Price (AED)"
    NDISP = len(V["ITEMS"]) * len(TR)
    for k in range(1, NDISP + 1):
        rr = 3 + k
        ws[f"T{rr}"] = k
        mi = f"MATCH({k},$I$4:$I${last},0)"
        mt = f"MATCH({k},$Q$4:$Q$8,0)"
        ws[f"U{rr}"] = (f'=IF({code}_Mode="Per equipment",IFERROR(INDEX($D$4:$D${last},{mi})&" ("&INDEX($A$4:$A${last},{mi})&")",""),'
                        f'IF({code}_Mode="Per unit (train)",IFERROR("Chemical cleaning – "&INDEX($O$4:$O$8,{mt}),""),'
                        f'IF({k}=1,"LUMP SUM – "&QProject,"")))')
        ws[f"V{rr}"] = (f'=IF({code}_Mode="Per equipment",IFERROR(INDEX($E$4:$E${last},{mi}),""),'
                        f'IF({code}_Mode="Per unit (train)",IFERROR(INDEX($O$4:$O$8,{mt}),""),IF({k}=1,"All units","")))')
        ws[f"W{rr}"] = (f'=IF({code}_Mode="Per equipment",IFERROR(INDEX($F$4:$F${last},{mi}),""),'
                        f'IF({code}_Mode="Per unit (train)",IFERROR(INDEX($P$4:$P$8,{mt}),""),IF({k}=1,$L$4,"")))')
        ws[f"X{rr}"] = (f'=IF({code}_Mode="Per equipment",IFERROR(INDEX($J$4:$J${last},{mi}),""),'
                        f'IF({code}_Mode="Per unit (train)",IFERROR(INDEX($R$4:$R$8,{mt}),""),IF({k}=1,{code}_Net,"")))')
    ws.sheet_state = "hidden"
    ws.ndisp = NDISP
    return ws


def build_client_quote(B, code, title, alloc, scope_text):
    ws = B.wb.create_sheet(f"{code} Client Quote")
    ws.sheet_view.showGridLines = False
    for c, w in {"A": 2, "B": 17, "C": 40, "D": 18, "E": 16, "F": 20, "G": 2}.items():
        ws.column_dimensions[c].width = w
    AQ = q(alloc.title)
    # letterhead
    ws.merge_cells("B1:F1"); ws["B1"] = "=UPPER(CoName)"; ws["B1"].font = Font(name=FN, size=18, bold=True, color=BLUE)
    ws.merge_cells("B2:F2"); ws["B2"] = "=CoTag"; ws["B2"].font = Font(name=FN, size=10, bold=True, color="FFFFFF")
    ws["B2"].fill = FILL_GREEN; ws["B2"].alignment = AL_LN
    for i, n in enumerate(["CoAddr", "CoTel", "CoWeb", "CoCert"]):
        ws.merge_cells(f"B{3+i}:F{3+i}"); ws[f"B{3+i}"] = f"={n}"; ws[f"B{3+i}"].font = Font(name=FN, size=9, color="404040")
    for c in "BCDEF":
        ws[f"{c}7"].border = Border(bottom=Side(style="thick", color=BLUE))
    ws.merge_cells("B9:F9"); ws["B9"] = f"QUOTATION – {title.upper()}"
    ws["B9"].font = Font(name=FN, size=14, bold=True, color=NAVY); ws["B9"].alignment = AL_C
    blank = lambda n: f'=IF({n}="","",{n})'
    info = [("Quotation No.", "=QNo&IF(QRev=\"\",\"\",\"  Rev \"&QRev)", "Date", blank("QDate")),
            ("Client", blank("QClient"), "Currency", "=QCurr"),
            ("Attention", blank("QAttn"), "Your ref.", blank("QRef")),
            ("Project / Plant", blank("QProject"), "Location", blank("QLoc"))]
    for i, (a, fa, b, fb) in enumerate(info):
        r = 11 + i
        ws[f"B{r}"] = a; ws[f"B{r}"].font = f_bold
        ws[f"C{r}"] = fa; ws[f"C{r}"].font = f_body; ws[f"C{r}"].alignment = AL_LN
        ws[f"E{r}"] = b; ws[f"E{r}"].font = f_bold; ws[f"E{r}"].alignment = AL_R
        ws[f"F{r}"] = fb; ws[f"F{r}"].font = f_body; ws[f"F{r}"].alignment = AL_LN
    ws["F11"].number_format = "dd-mmm-yyyy"
    ws.merge_cells("B16:F16"); ws["B16"] = scope_text; ws["B16"].font = f_body; ws["B16"].alignment = AL_L
    ws.row_dimensions[16].height = 30
    # price summary
    section_bar(ws, 18, "PRICE SUMMARY", "B", "F", fill=FILL_HDR)
    cur = '"("&QCurr&")"'
    summ = [("Price before discount", f"={code}_PriceGross/QRate"),
            ("Discount", f"={code}_Discount/QRate"),
            ("TOTAL PRICE EXCL. VAT", f"={code}_Net/QRate"),
            ('="VAT @ "&TEXT(QVAT,"0%")', f"={code}_Net/QRate*N(QVAT)"),
            ("TOTAL PRICE INCL. VAT", f"={code}_Net/QRate*(1+N(QVAT))")]
    for i, (lab, f) in enumerate(summ):
        r = 19 + i
        strong = lab.startswith("TOTAL")
        ws.merge_cells(f"B{r}:D{r}")
        ws[f"B{r}"] = lab if lab.startswith("=") else f'="{lab} "&{cur}'
        ws[f"B{r}"].font = f_bold if strong else f_body; ws[f"B{r}"].alignment = AL_R
        ws.merge_cells(f"E{r}:F{r}")
        ws[f"E{r}"] = f; style(ws[f"E{r}"], f_bold if strong else f_body, FILL_KEY if strong else None, nf=NF_AED, al=AL_R)
        if strong:
            for c in "BCD":
                ws[f"{c}{r}"].fill = FILL_KEY
    # terms
    section_bar(ws, 25, "COMMERCIAL TERMS", "B", "F", fill=FILL_HDR)
    terms = [("Validity", "QTValidity"), ("Payment terms", "QTPayment"), ("Mobilisation", "QTMob"),
             ("Exclusions", "QTExcl"), ("By client", "QTClient"), ("Notes", "QTNotes")]
    for i, (t, n) in enumerate(terms):
        r = 26 + i
        ws[f"B{r}"] = t; ws[f"B{r}"].font = f_bold; ws[f"B{r}"].alignment = AL_LN
        ws.merge_cells(f"C{r}:F{r}"); ws[f"C{r}"] = f'=IF({n}="","–",{n})'; ws[f"C{r}"].alignment = AL_L
        ws[f"C{r}"].font = f_body; ws.row_dimensions[r].height = 28
    r = 26 + len(terms) + 1
    ws[f"B{r}"] = "For and on behalf of"; ws[f"B{r}"].font = f_body
    ws[f"B{r+1}"] = "=CoName"; ws[f"B{r+1}"].font = f_bold
    ws[f"B{r+3}"] = '=IF(QPrep="","",QPrep)'; ws[f"B{r+3}"].font = f_body
    # breakdown
    r0 = r + 6
    section_bar(ws, r0, "PRICE BREAKDOWN", "B", "F", fill=FILL_HDR)
    for j, h in enumerate(["Item", "Description", "Unit / train", "Volume (m³)", '="Price ("&QCurr&")"']):
        style(ws.cell(row=r0 + 1, column=2 + j, value=h), f_hdr, FILL_SECT, al=AL_C)
    for k in range(1, alloc.ndisp + 1):
        r = r0 + 1 + k
        a = 3 + k
        ws[f"B{r}"] = f'=IF({AQ}!U{a}="","",{k})'
        ws[f"C{r}"] = f"={AQ}!U{a}"
        ws[f"D{r}"] = f"={AQ}!V{a}"
        ws[f"E{r}"] = f"={AQ}!W{a}"
        ws[f"F{r}"] = f'=IF({AQ}!X{a}="","",{AQ}!X{a}/QRate)'
        for c, nf in (("B", "0"), ("C", None), ("D", None), ("E", NF_M3), ("F", NF_AED)):
            ws[f"{c}{r}"].font = f_body; ws[f"{c}{r}"].alignment = AL_C if c != "C" else AL_LN
            if nf: ws[f"{c}{r}"].number_format = nf
    last = r0 + 1 + alloc.ndisp
    ws.conditional_formatting.add(f"B{r0+2}:F{last}", FormulaRule(formula=[f'$C{r0+2}<>""'], border=BORDER))
    rt = r0 + 2
    ws[f"H{r0}"] = "Rows used"; ws[f"H{r0}"].font = f_grey
    ws[f"I{r0}"] = f"={code}_NItems"; ws[f"I{r0}"].font = f_grey
    ws[f"H{r0+1}"] = "Print up to row"; ws[f"H{r0+1}"].font = f_grey
    ws[f"I{r0+1}"] = f"={r0+1}+{code}_NItems"; ws[f"I{r0+1}"].font = f_grey
    idx = B.wb.sheetnames.index(ws.title)
    ws.page_setup.orientation = "portrait"; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.dyn_print = (idx, r0 + 1)
    return ws


# ================================================================ COMBINED QUOTATION (master only)
def build_quotation(B, built):
    ws = B.wb.create_sheet("Quotation", 2)
    ws.sheet_view.showGridLines = False
    for c, w in {"A": 2, "B": 6, "C": 40, "D": 14, "E": 20, "F": 20, "G": 40}.items():
        ws.column_dimensions[c].width = w
    ws["B1"] = "COMBINED QUOTATION – multi-service jobs"; ws["B1"].font = f_title
    ws["B2"] = "Tick the services included in this quotation. Prices come from each service's costing sheet."; ws["B2"].font = f_sub
    for j, h in enumerate(["#", "Service", "Include?", "Net price (AED)", '="Net price ("&QCurr&")"', "Status"]):
        style(ws.cell(row=4, column=2 + j, value=h), f_hdr, FILL_HDR, al=AL_C)
    for i, (code, nm, ok) in enumerate(SERVICES):
        r = 5 + i
        ws[f"B{r}"] = i + 1; style(ws[f"B{r}"], f_grey, al=AL_C)
        label(ws, f"C{r}", nm, f_body)
        inp(ws, f"D{r}", "Yes" if code in built else "No"); B.dv(ws, '"Yes,No"', f"D{r}")
        if code in built:
            ws[f"E{r}"] = f'=IF(D{r}="Yes",{code}_Net,0)'; ws[f"G{r}"] = "Template built"
        else:
            ws[f"E{r}"] = 0; ws[f"G{r}"] = "Template not built yet (pilot)"
        style(ws[f"E{r}"], f_body, FILL_RES, nf=NF_AED, al=AL_C)
        ws[f"F{r}"] = f"=E{r}/QRate"; style(ws[f"F{r}"], f_body, FILL_RES, nf=NF_AED, al=AL_C)
        style(ws[f"G{r}"], f_grey, al=AL_LN)
    r = 5 + len(SERVICES)
    rows = [("TOTAL EXCL. VAT", f"=SUM(E5:E{r-1})"), ('="VAT @ "&TEXT(QVAT,"0%")', f"=E{r}*N(QVAT)"),
            ("TOTAL INCL. VAT", f"=E{r}+E{r+1}")]
    for i, (lab, f) in enumerate(rows):
        x = r + i
        label(ws, f"C{x}", lab)
        ws[f"E{x}"] = f; ws[f"F{x}"] = f"=E{x}/QRate"
        for c in "EF":
            style(ws[f"{c}{x}"], f_bold, FILL_KEY, nf=NF_AED, al=AL_C)
    return ws


def finish(B, out):
    # dynamic print areas for client quotes (letterhead + used breakdown rows)
    for ws in B.wb.worksheets:
        if hasattr(ws, "dyn_print"):
            _, hdr_row = ws.dyn_print
            code = ws.title.split()[0]
            dn = DefinedName("_xlnm.Print_Area", attr_text=f"OFFSET({q(ws.title)}!$A$1,0,0,{hdr_row}+MAX(1,{code}_NItems),7)")
            ws.defined_names["_xlnm.Print_Area"] = dn
    B.wb.save(out)
    print("saved", out)


def build(services, master, out):
    B = Book()
    build_quote_info(B)
    build_rates(B)
    built = []
    for code in services:
        if code == "CC":
            build_cc(B)
            built.append(code)
    if master:
        build_quotation(B, built)
    # sheet order: inputs first
    order = ["Quote Info", "Rates", "Quotation"]
    for code in built:
        order += [f"{code} Vol Settings", f"{code} Volume Calc", f"{code} Costing", f"{code} Client Quote"]
    sheets = B.wb._sheets
    B.wb._sheets = [s for n in order for s in sheets if s.title == n] + [s for s in sheets if s.title not in order]
    B.wb.active = 0
    finish(B, out)


if __name__ == "__main__":
    outdir = sys.argv[1] if len(sys.argv) > 1 else "pricing"
    os.makedirs(outdir, exist_ok=True)
    build(["CC"], True, os.path.join(outdir, "Delight_Pricing_Master.xlsx"))
    build(["CC"], False, os.path.join(outdir, "Delight_Pricing_Chemical_Cleaning.xlsx"))
