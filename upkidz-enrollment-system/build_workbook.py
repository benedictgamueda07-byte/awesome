"""
Builds UpKidz_Enrollment_System.xlsx — the workbook half of the UpKidz
Enrollment Management System. The VBA half lives in UpKidzEnrollment.bas.

Run:  python3 build_workbook.py [output_path]

Layout constants here MUST match the constants at the top of the .bas module
(FIRST_ROW, LAST_ROW, section start columns, defined names).
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference

OUT = sys.argv[1] if len(sys.argv) > 1 else "UpKidz_Enrollment_System.xlsx"

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
PROGRAMS = ["Toddler PlayTrack", "Preschool PlayTrack", "One-on-One Tutoring"]

FIRST_ROW, LAST_ROW = 7, 106          # 100 enrollment rows per section
FONT = "Arial"
PESO = '"₱"#,##0.00;-"₱"#,##0.00;"-"'

# ---- section definitions (start column = the "No." column) ----------------
SECTIONS = [
    dict(key="Toddler", title="TODDLERS GROUP", program="Toddler PlayTrack", start=2,
         cols=["No.", "Full Name", "Schedule", "Package", "Discount", "Price"],
         dark="2E7D32", mid="A5D6A7", light="E8F5E9", tab="2E7D32"),
    dict(key="Preschool", title="PRESCHOOL GROUP", program="Preschool PlayTrack", start=10,
         cols=["No.", "Full Name", "Schedule", "Package", "Discount", "Price"],
         dark="1E88E5", mid="90CAF9", light="E3F2FD", tab="1E88E5"),
    dict(key="Tutor", title="ONE-ON-ONE TUTORING", program="One-on-One Tutoring", start=18,
         cols=["No.", "Full Name", "Schedule", "Package", "Price"],
         dark="1A237E", mid="9FA8DA", light="E8EAF6", tab="1A237E"),
]
WIDTHS = {"No.": 5.5, "Full Name": 27, "Schedule": 32, "Package": 12, "Discount": 11.5, "Price": 13}

BANNER = "1F2A44"
ACCENT = "FFC107"
GREY_TXT = "6B7280"
INPUT_FILL = "FFFDE7"

thin = Side(style="thin", color="B0BEC5")
med = Side(style="medium", color="455A64")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def font(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", indent=1)
RIGHT = Alignment(horizontal="right", vertical="center")

wb = Workbook()

# =============================================================================
# SETTINGS (lists, packages, configuration)
# =============================================================================
st = wb.active
st.title = "Settings"
st.sheet_properties.tabColor = "607D8B"
st.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJ", [24, 16, 10, 16, 34, 3, 16, 3, 36, 3]):
    st.column_dimensions[col].width = w

st["A1"] = "SYSTEM SETTINGS & LISTS"
st["A1"].font = font(14, True, "FFFFFF")
st.merge_cells("A1:I1")
st["A1"].fill = fill(BANNER)
st["A1"].alignment = LEFT
st.row_dimensions[1].height = 28

cfg = [("School Year", 2026, "Shown in monthly sheet titles."),
       ("Price per month", "Full package price",
        "Full package price = record the full net price in every covered month "
        "(as in the spec example). Split per month = net price ÷ duration in each month.")]
for i, (k, v, note) in enumerate(cfg, start=2):
    st.cell(i, 1, k).font = font(10, True)
    c = st.cell(i, 2, v)
    c.font = font(10, True, "0000FF")
    c.fill = fill(INPUT_FILL)
    c.border = BORDER
    c.alignment = CENTER
    st.cell(i, 3, note).font = font(9, color=GREY_TXT, italic=True)
dv_mode = DataValidation(type="list", formula1='"Full package price,Split per month"', allow_blank=False)
st.add_data_validation(dv_mode)
dv_mode.add("B3")

# Package table A5:E
st["A5"], st["B5"], st["C5"], st["D5"], st["E5"] = "Program", "Package", "Months", "Standard Fee", "Key (auto)"
for c in "ABCDE":
    st[f"{c}5"].font = font(10, True, "FFFFFF")
    st[f"{c}5"].fill = fill(BANNER)
    st[f"{c}5"].alignment = CENTER
    st[f"{c}5"].border = BORDER
packages = [
    ("Toddler PlayTrack", "1 Month", 1, 3500), ("Toddler PlayTrack", "3 Months", 3, 9900),
    ("Toddler PlayTrack", "6 Months", 6, 18700), ("Toddler PlayTrack", "12 Months", 12, 35000),
    ("Preschool PlayTrack", "1 Month", 1, 3800), ("Preschool PlayTrack", "3 Months", 3, 10800),
    ("Preschool PlayTrack", "6 Months", 6, 20500), ("Preschool PlayTrack", "12 Months", 12, 38000),
    ("One-on-One Tutoring", "1 Month", 1, 4500), ("One-on-One Tutoring", "3 Months", 3, 12900),
    ("One-on-One Tutoring", "6 Months", 6, 24500), ("One-on-One Tutoring", "12 Months", 12, 46000),
]
PKG_FIRST, PKG_LAST = 6, 40           # room for more packages
for r in range(PKG_FIRST, PKG_LAST + 1):
    i = r - PKG_FIRST
    if i < len(packages):
        for j, v in enumerate(packages[i]):
            st.cell(r, j + 1, v)
    for j in range(1, 5):
        c = st.cell(r, j)
        c.font = font(10, color="0000FF")
        c.border = BORDER
        c.fill = fill(INPUT_FILL)
    st.cell(r, 3).alignment = CENTER
    st.cell(r, 4).number_format = '"₱"#,##0'
    k = st.cell(r, 5, f'=IF(A{r}="","",A{r}&"|"&B{r})')
    k.font = font(9, color=GREY_TXT)
    k.border = BORDER
st["D8"].comment = Comment("₱18,700 for Toddler 6 Months is from the client's example. "
                           "All other fees are SAMPLE placeholders — replace with actual UpKidz rates.",
                           "UpKidz System")
st.cell(PKG_LAST + 2, 1, "Keep rows grouped by program (all Toddler rows together, etc.). "
        "The Package dropdown on the Dashboard reads this table. Fees other than "
        "Toddler 6 Months (₱18,700, from client example) are sample values — edit them.").font = font(9, color=GREY_TXT, italic=True)

# Month list G5:G17
st["G5"] = "Months"
st["G5"].font = font(10, True, "FFFFFF")
st["G5"].fill = fill(BANNER)
st["G5"].alignment = CENTER
for i, m in enumerate(MONTHS):
    c = st.cell(6 + i, 7, m)
    c.border = BORDER
    c.font = font(10)

# Programs list G20:G22
st["G20"] = "Programs"
st["G20"].font = font(10, True, "FFFFFF")
st["G20"].fill = fill(BANNER)
st["G20"].alignment = CENTER
for i, p in enumerate(PROGRAMS):
    c = st.cell(21 + i, 7, p)
    c.border = BORDER
    c.font = font(10)

# Schedules list I5:I30 (free text still allowed on dashboard)
st["I5"] = "Schedules (dropdown suggestions)"
st["I5"].font = font(10, True, "FFFFFF")
st["I5"].fill = fill(BANNER)
st["I5"].alignment = CENTER
scheds = ["Monday and Wednesday, 9:00–10:00 AM", "Tuesday and Thursday, 9:00–10:00 AM",
          "Monday and Wednesday, 10:30–11:30 AM", "Tuesday and Thursday, 10:30–11:30 AM",
          "Monday, Wednesday and Friday, 9:00–10:30 AM", "Saturday, 9:00–11:00 AM",
          "Monday to Friday, 1:00–2:00 PM", "Saturday, 1:00–2:00 PM"]
for r in range(6, 31):
    c = st.cell(r, 9, scheds[r - 6] if r - 6 < len(scheds) else None)
    c.border = BORDER
    c.font = font(10, color="0000FF")
    c.fill = fill(INPUT_FILL)

st.freeze_panes = "A6"

dn = wb.defined_names
def name(n, ref):
    dn[n] = DefinedName(n, attr_text=ref)

name("MonthList", "Settings!$G$6:$G$17")
name("ProgramList", "Settings!$G$21:$G$23")
name("ScheduleList", f"Settings!$I$6:$I$30")
name("PkgProgram", f"Settings!$A${PKG_FIRST}:$A${PKG_LAST}")
name("PkgName", f"Settings!$B${PKG_FIRST}:$B${PKG_LAST}")
name("PkgMonths", f"Settings!$C${PKG_FIRST}:$C${PKG_LAST}")
name("PkgFee", f"Settings!$D${PKG_FIRST}:$D${PKG_LAST}")
name("PkgKey", f"Settings!$E${PKG_FIRST}:$E${PKG_LAST}")
name("cfgSchoolYear", "Settings!$B$2")
name("cfgPriceMode", "Settings!$B$3")
name("PackageChoices",
     f"OFFSET(Settings!$B${PKG_FIRST},MATCH(Dashboard!$C$6,Settings!$A${PKG_FIRST}:$A${PKG_LAST},0)-1,0,"
     f"MAX(1,COUNTIF(Settings!$A${PKG_FIRST}:$A${PKG_LAST},Dashboard!$C$6)),1)")

# =============================================================================
# DASHBOARD
# =============================================================================
ds = wb.create_sheet("Dashboard", 0)
ds.sheet_properties.tabColor = ACCENT
ds.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJK", [2, 22, 22, 22, 34, 3, 18, 18, 18, 18, 2]):
    ds.column_dimensions[col].width = w

def banner(ws, rng1, rng2, title, sub):
    ws.merge_cells(rng1)
    ws.merge_cells(rng2)
    a = ws[rng1.split(":")[0]]
    a.value = title
    a.font = font(20, True, ACCENT)
    a.fill = fill(BANNER)
    a.alignment = Alignment(horizontal="center", vertical="center")
    b = ws[rng2.split(":")[0]]
    b.value = sub
    b.font = font(11, False, "FFFFFF")
    b.fill = fill(BANNER)
    b.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws[rng1] + ws[rng2]:
        for c in row:
            c.fill = fill(BANNER)

banner(ds, "B1:J1", "B2:J2", "UPKIDZ PHILIPPINES",
       "Enrollment Management System  •  Dashboard")
ds.row_dimensions[1].height = 36
ds.row_dimensions[2].height = 20
ds.row_dimensions[3].height = 10

def card_header(ws, rng, text, color):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = text
    c.font = font(12, True, "FFFFFF")
    c.alignment = LEFT
    for row in ws[rng]:
        for x in row:
            x.fill = fill(color)

card_header(ds, "B4:E4", "NEW ENROLLMENT", "37474F")
ds.row_dimensions[4].height = 24
ds["B5"] = "Mode"
ds["B5"].font = font(9, True, GREY_TXT)
ds["C5"] = None   # editing enrollment ID (set by VBA). Blank = new enrollment
ds["C5"].font = font(10, True, "C62828")
ds["D5"] = '=IF(C5="","New enrollment — fill in the yellow cells","EDIT MODE: changes will be applied with UPDATE ENROLLMENT")'
ds.merge_cells("D5:E5")
ds["D5"].font = font(9, False, GREY_TXT, True)

fields = [  # row, label, name, kind
    (6, "Program", "inProgram", "input"),
    (7, "Full Name", "inName", "input"),
    (8, "Schedule", "inSchedule", "input"),
    (9, "Package", "inPackage", "input"),
    (10, "Start Month", "inStartMonth", "input"),
    (11, "Duration (Months)", "inDuration", "input"),
    (12, "Program Fee (₱)", "inFee", "input"),
    (13, "Discount (₱)", "inDiscount", "input"),
    (14, "Duration Display", "outDurationDisplay", "calc"),
    (15, "Net Price (₱)", "outNetPrice", "calc"),
    (16, "Status Check", "outStatus", "calc"),
]
for r, label, nm, kind in fields:
    ds.row_dimensions[r].height = 22
    lc = ds.cell(r, 2, label)
    lc.font = font(10, True, "263238")
    lc.alignment = LEFT
    lc.fill = fill("ECEFF1")
    lc.border = BORDER
    ds.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    for cc in (3, 4):
        x = ds.cell(r, cc)
        x.border = BORDER
        x.fill = fill(INPUT_FILL if kind == "input" else "F5F5F5")
    v = ds.cell(r, 3)
    v.font = font(11, kind == "calc", "0000FF" if kind == "input" else "000000")
    v.alignment = LEFT
    name(nm, f"Dashboard!$C${r}")
name("inEditID", "Dashboard!$C$5")

ds["C12"].number_format = PESO
ds["C13"].number_format = PESO
ds["C15"].number_format = PESO
ds["C15"].font = font(13, True, "1B5E20")
ds["C11"].alignment = LEFT

ds["C14"] = ('=IF(OR(C10="",C11=""),"",IF(ISNA(MATCH(C10,MonthList,0)),"Invalid start month",'
             'IF(OR(NOT(ISNUMBER(C11)),C11<1,C11>12),"Duration must be 1–12",'
             'C10&IF(C11>1,"–"&INDEX(MonthList,MOD(MATCH(C10,MonthList,0)+C11-2,12)+1),"")'
             '&" ("&C11&IF(C11=1," month)"," months)"))))')
ds["C15"] = '=IF(C12="","",C12-N(C13))'
ds["C16"] = ('=IF(C6="","⚠ Select a program",IF(TRIM(C7)="","⚠ Enter the full name",'
             'IF(TRIM(C8)="","⚠ Enter the schedule",IF(C9="","⚠ Select a package",'
             'IF(C10="","⚠ Select the start month",IF(OR(NOT(ISNUMBER(C11)),C11<1,C11>12),"⚠ Enter duration (1–12)",'
             'IF(OR(NOT(ISNUMBER(C12)),C12<=0),"⚠ Enter the program fee",'
             'IF(N(C13)>C12,"⚠ Discount exceeds the program fee","✔ Ready to submit"))))))))')
ds["C16"].font = font(11, True, "000000")

# hints (column E)
hints = {
    6: "Toddler / Preschool / One-on-One",
    7: "Student's complete name",
    8: "Pick from list or type any schedule",
    9: ('=IFERROR("Package covers "&INDEX(PkgMonths,MATCH(C6&"|"&C9,PkgKey,0))&" month(s)",'
        'IF(C6="","Choose a program first","Choose a package"))'),
    10: "January – December",
    11: ('=IFERROR(IF(AND(ISNUMBER(C11),C11<>INDEX(PkgMonths,MATCH(C6&"|"&C9,PkgKey,0))),'
         '"⚠ Differs from package length","Months covered by enrollment"),"Months covered by enrollment")'),
    12: ('=IFERROR("Standard fee: ₱"&TEXT(INDEX(PkgFee,MATCH(C6&"|"&C9,PkgKey,0)),"#,##0"),'
         '"Amount before discount")'),
    13: '=IF(AND(ISNUMBER(C12),N(C13)>C12),"⚠ Discount is greater than fee!","Leave blank if none")',
    14: "Automatic",
    15: "Automatic: Program Fee − Discount",
    16: "Must read ✔ before submitting",
}
for r, h in hints.items():
    c = ds.cell(r, 5, h)
    c.font = font(9, False, GREY_TXT, True)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)

# status colouring
ds.conditional_formatting.add("C16", FormulaRule(formula=['LEFT($C$16,1)="✔"'],
                              fill=fill("C8E6C9"), font=Font(name=FONT, bold=True, color="1B5E20")))
ds.conditional_formatting.add("C16", FormulaRule(formula=['LEFT($C$16,1)="⚠"'],
                              fill=fill("FFE0B2"), font=Font(name=FONT, bold=True, color="BF360C")))
ds.conditional_formatting.add("E9:E13", FormulaRule(formula=['LEFT(E9,1)="⚠"'],
                              font=Font(name=FONT, bold=True, color="C62828")))

# data validation
dv = DataValidation(type="list", formula1="=ProgramList", allow_blank=False,
                    error="Choose Toddler PlayTrack, Preschool PlayTrack or One-on-One Tutoring.",
                    errorTitle="Invalid program")
dv.showErrorMessage = True
ds.add_data_validation(dv); dv.add("C6")
dv = DataValidation(type="list", formula1="=ScheduleList", allow_blank=True)
dv.showErrorMessage = False            # suggestions only; free text allowed
ds.add_data_validation(dv); dv.add("C8")
dv = DataValidation(type="list", formula1="=PackageChoices", allow_blank=False,
                    error="Choose a package listed for the selected program (see Settings).",
                    errorTitle="Invalid package")
dv.showErrorMessage = True
ds.add_data_validation(dv); dv.add("C9")
dv = DataValidation(type="list", formula1="=MonthList", allow_blank=False,
                    error="Choose a month from January to December.", errorTitle="Invalid month")
dv.showErrorMessage = True
ds.add_data_validation(dv); dv.add("C10")
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="12",
                    error="Duration must be a whole number from 1 to 12.", errorTitle="Invalid duration")
dv.showErrorMessage = True
ds.add_data_validation(dv); dv.add("C11")
dv = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                    error="Program fee must be a number greater than zero.", errorTitle="Invalid fee")
dv.showErrorMessage = True
ds.add_data_validation(dv); dv.add("C12")
dv = DataValidation(type="custom", formula1="AND(ISNUMBER(C13),C13>=0,C13<=N(C12))",
                    error="Discount must be zero or more and cannot exceed the Program Fee.",
                    errorTitle="Invalid discount", errorStyle="warning")
dv.showErrorMessage = True
ds.add_data_validation(dv); dv.add("C13")

# button zone
ds.row_dimensions[17].height = 8
ds.row_dimensions[18].height = 34
ds.row_dimensions[19].height = 8
ds["B18"] = "Buttons appear here after running the macro  InstallUpKidzSystem  (see README)."
ds.merge_cells("B18:E18")
ds["B18"].font = font(9, False, GREY_TXT, True)
ds["B18"].alignment = CENTER

# ---- right column: manage + snapshot ---------------------------------------
card_header(ds, "G4:J4", "MANAGE ENROLLMENTS", "37474F")
def labeled_input(r, label, nm):
    ds.row_dimensions[r].height = 22
    c = ds.cell(r, 7, label)
    c.font = font(10, True, "263238"); c.fill = fill("ECEFF1"); c.border = BORDER; c.alignment = LEFT
    ds.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
    for cc in (8, 9, 10):
        x = ds.cell(r, cc); x.fill = fill(INPUT_FILL); x.border = BORDER
    ds.cell(r, 8).font = font(11, False, "0000FF"); ds.cell(r, 8).alignment = LEFT
    name(nm, f"Dashboard!$H${r}")
labeled_input(6, "Enrollment ID", "inManageID")
ds["G7"] = "Load an ID to edit it, or cancel it. IDs appear in Enrollment History and search results."
ds.merge_cells("G7:J7"); ds["G7"].font = font(9, False, GREY_TXT, True)
ds.row_dimensions[8].height = 30          # LOAD / UPDATE / CANCEL buttons
labeled_input(10, "Search Student", "inSearch")
ds["G11"] = "Type all or part of a name. Results open on the Student Search sheet."
ds.merge_cells("G11:J11"); ds["G11"].font = font(9, False, GREY_TXT, True)
ds.row_dimensions[12].height = 30          # SEARCH button

card_header(ds, "G14:J14", "MONTHLY SNAPSHOT", "37474F")
ds["G15"] = "Month"; ds["G15"].font = font(10, True, "263238"); ds["G15"].fill = fill("ECEFF1")
ds["G15"].border = BORDER; ds["G15"].alignment = LEFT
ds.merge_cells("H15:J15")
ds["H15"] = "January"
for cc in (8, 9, 10):
    ds.cell(15, cc).fill = fill(INPUT_FILL); ds.cell(15, cc).border = BORDER
ds["H15"].font = font(11, True, "0000FF"); ds["H15"].alignment = LEFT
dv = DataValidation(type="list", formula1="=MonthList", allow_blank=False)
ds.add_data_validation(dv); dv.add("H15")
ds["G16"], ds["H16"], ds["I16"] = "Program", "Students", "Total Value"
ds.merge_cells("I16:J16")
for cc in (7, 8, 9, 10):
    c = ds.cell(16, cc); c.font = font(9, True, "FFFFFF"); c.fill = fill("546E7A")
    c.alignment = CENTER; c.border = BORDER
snap = [("Toddlers Group", "G", SECTIONS[0]), ("Preschool Group", "O", SECTIONS[1]),
        ("One-on-One", "V", SECTIONS[2])]
for i, (lbl, pcol, sec) in enumerate(snap):
    r = 17 + i
    c = ds.cell(r, 7, lbl); c.font = font(10, True, "FFFFFF"); c.fill = fill(sec["dark"]); c.border = BORDER
    c.alignment = LEFT
    c = ds.cell(r, 8, f"=IFERROR(INDIRECT(\"'\"&$H$15&\"'!{pcol}5\"),0)")
    c.font = font(11, True); c.alignment = CENTER; c.border = BORDER
    ds.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    c = ds.cell(r, 9, f"=IFERROR(INDIRECT(\"'\"&$H$15&\"'!{pcol}4\"),0)")
    c.font = font(11, True); c.number_format = PESO; c.alignment = RIGHT; c.border = BORDER
    ds.cell(r, 10).border = BORDER
r = 20
c = ds.cell(r, 7, "TOTAL"); c.font = font(10, True); c.fill = fill("CFD8DC"); c.border = BORDER; c.alignment = LEFT
c = ds.cell(r, 8, "=SUM(H17:H19)"); c.font = font(11, True); c.alignment = CENTER; c.border = BORDER; c.fill = fill("CFD8DC")
ds.merge_cells("I20:J20")
c = ds.cell(r, 9, "=SUM(I17:I19)"); c.font = font(11, True); c.number_format = PESO; c.alignment = RIGHT
c.border = BORDER; c.fill = fill("CFD8DC"); ds.cell(r, 10).border = BORDER; ds.cell(r, 10).fill = fill("CFD8DC")
for rr in (17, 18, 19, 20):
    ds.row_dimensions[rr].height = 20

ds["G21"] = "=\"Active enrollments on record: \"&COUNTIF('Enrollment History'!$M:$M,\"Active\")"
ds.merge_cells("G21:J21"); ds["G21"].font = font(9, True, GREY_TXT)

# How-to legend
ds["B21"] = "HOW TO USE"
ds["B21"].font = font(10, True, "263238")
steps = [
    "1. Fill in the yellow cells (blue text = your inputs). Grey cells calculate automatically.",
    "2. Make sure Status Check shows ✔ Ready to submit, then click SUBMIT ENROLLMENT.",
    "3. The student is copied to the correct program section of every covered month (Nov + 3 months → Nov, Dec, Jan).",
    "4. To edit: type the Enrollment ID → LOAD FOR EDIT → change fields → UPDATE ENROLLMENT.",
    "5. To cancel: type the Enrollment ID → CANCEL ENROLLMENT (removed from every month, kept in history).",
    "Example: Toddler PlayTrack · Benedict Gamueda · Monday and Wednesday, 9:00–10:00 AM · 6 Months · January · 6 · ₱18,700 · ₱2,000",
]
for i, s in enumerate(steps):
    ds.merge_cells(start_row=22 + i, start_column=2, end_row=22 + i, end_column=5)
    c = ds.cell(22 + i, 2, s)
    c.font = font(9, False, "37474F", i == len(steps) - 1)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ds.row_dimensions[22 + i].height = 18

ds.page_setup.orientation = "landscape"
ds.page_setup.fitToWidth = 1
ds.page_setup.fitToHeight = 1
ds.sheet_properties.pageSetUpPr.fitToPage = True
ds.print_area = "A1:K28"

# =============================================================================
# MONTHLY SHEETS
# =============================================================================
def build_month(ws, month):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.column_dimensions["A"].width = 1.5
    last_col = SECTIONS[-1]["start"] + len(SECTIONS[-1]["cols"])  # ID col of last section
    # banner B1:V2
    endc = L(last_col - 1)
    ws.merge_cells(f"B1:{endc}1")
    ws.merge_cells(f"B2:{endc}2")
    ws["B1"] = "UPKIDZ PHILIPPINES"
    ws["B1"].font = font(18, True, ACCENT)
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B2"] = f'="{month.upper()} "&cfgSchoolYear&"  •  MONTHLY ENROLLMENT RECORD"'
    ws["B2"].font = font(11, True, "FFFFFF")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    for rr in (1, 2):
        for cc in range(2, last_col):
            ws.cell(rr, cc).fill = fill(BANNER)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 8
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 22

    for sec in SECTIONS:
        s = sec["start"]
        n = len(sec["cols"])
        e = s + n - 1               # last visible column (Price)
        idc = e + 1                 # hidden Enrollment ID column
        name_c, price_c = L(s + 1), L(e)
        for j, h in enumerate(sec["cols"]):
            ws.column_dimensions[L(s + j)].width = WIDTHS[h]
        ws.column_dimensions[L(idc)].width = 12
        ws.column_dimensions[L(idc)].hidden = True
        if idc + 1 < SECTIONS[-1]["start"] + 10 and sec is not SECTIONS[-1]:
            ws.column_dimensions[L(idc + 1)].width = 2.5   # spacer

        # row 4: title (left) + TOTAL (top-right corner)
        ws.merge_cells(start_row=4, start_column=s, end_row=4, end_column=e - 2)
        t = ws.cell(4, s, sec["title"])
        t.font = font(13, True, "FFFFFF")
        t.alignment = LEFT
        lab = ws.cell(4, e - 1, "TOTAL")
        lab.font = font(10, True, "FFFFFF")
        lab.alignment = RIGHT
        tot = ws.cell(4, e, f"=SUM({price_c}{FIRST_ROW}:{price_c}{LAST_ROW})")
        tot.font = font(11, True, sec["dark"])
        tot.fill = fill("FFFFFF")
        tot.number_format = PESO
        tot.alignment = RIGHT
        tot.border = Border(left=med, right=med, top=med, bottom=med)
        for cc in range(s, e):
            ws.cell(4, cc).fill = fill(sec["dark"])
        # row 5: program name + STUDENTS count
        ws.merge_cells(start_row=5, start_column=s, end_row=5, end_column=e - 2)
        p = ws.cell(5, s, sec["program"])
        p.font = font(9, True, sec["dark"], True)
        p.alignment = LEFT
        lab = ws.cell(5, e - 1, "STUDENTS")
        lab.font = font(8, True, sec["dark"])
        lab.alignment = RIGHT
        cnt = ws.cell(5, e, f"=COUNTA({name_c}{FIRST_ROW}:{name_c}{LAST_ROW})")
        cnt.font = font(10, True, sec["dark"])
        cnt.alignment = RIGHT
        for cc in range(s, e + 1):
            ws.cell(5, cc).fill = fill(sec["mid"])
        # row 6: headers
        for j, h in enumerate(sec["cols"]):
            c = ws.cell(6, s + j, h)
            c.font = font(10, True, "FFFFFF")
            c.fill = fill(sec["dark"])
            c.alignment = CENTER
            c.border = Border(left=Side(style="thin", color="FFFFFF"),
                              right=Side(style="thin", color="FFFFFF"),
                              bottom=Side(style="medium", color=sec["dark"]))
        hid = ws.cell(6, idc, "Enrollment ID")
        hid.font = font(8, True, GREY_TXT)
        # data rows
        for r in range(FIRST_ROW, LAST_ROW + 1):
            band = fill(sec["light"]) if (r - FIRST_ROW) % 2 else fill("FFFFFF")
            for j, h in enumerate(sec["cols"]):
                c = ws.cell(r, s + j)
                c.fill = band
                c.border = BORDER
                c.font = font(10)
                if h == "No.":
                    c.value = f'=IF({name_c}{r}="","",COUNTA({name_c}${FIRST_ROW}:{name_c}{r}))'
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.font = font(10, True, sec["dark"])
                elif h in ("Discount", "Price"):
                    c.number_format = PESO
                    c.alignment = RIGHT
                elif h == "Package":
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.cell(r, idc).font = font(8, color=GREY_TXT)
            ws.row_dimensions[r].height = 17
        # bottom outline
        for cc in range(s, e + 1):
            c = ws.cell(LAST_ROW, cc)
            c.border = Border(left=thin, right=thin, top=thin, bottom=Side(style="medium", color=sec["dark"]))

    foot = ws.cell(LAST_ROW + 2, 2, "100 enrollment slots per program. Records are added automatically from the "
                   "Dashboard — use the Dashboard (Edit / Cancel) to change them so all months stay consistent.")
    foot.font = font(8, False, GREY_TXT, True)

    ws.freeze_panes = f"B{FIRST_ROW}"
    ws.print_title_rows = "1:6"
    ws.print_area = f"A1:{endc}{LAST_ROW}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = f"UpKidz Philippines — {month} — Page &P of &N"


for i, m in enumerate(MONTHS):
    ws = wb.create_sheet(m)
    ws.sheet_properties.tabColor = ["4CAF50", "2196F3", "3F51B5"][i % 3]
    build_month(ws, m)

# =============================================================================
# ENROLLMENT HISTORY (log, written by VBA)
# =============================================================================
lg = wb.create_sheet("Enrollment History")
lg.sheet_properties.tabColor = "8D6E63"
lg.sheet_view.showGridLines = False
hdr = ["Enrollment ID", "Date Submitted", "Program", "Full Name", "Schedule", "Package",
       "Start Month", "Duration", "Months Covered", "Program Fee", "Discount", "Net Price",
       "Status", "Last Updated", "Remarks"]
widths = [14, 17, 20, 26, 30, 11, 12, 9, 34, 13, 12, 13, 11, 17, 34]
banner(lg, "A1:O1", "A2:O2", "ENROLLMENT HISTORY",
       "Every submission, edit and cancellation is tracked here automatically — do not edit IDs")
lg.row_dimensions[1].height = 30
for j, (h, w) in enumerate(zip(hdr, widths)):
    lg.column_dimensions[L(j + 1)].width = w
    c = lg.cell(4, j + 1, h)
    c.font = font(10, True, "FFFFFF"); c.fill = fill("37474F"); c.alignment = CENTER; c.border = BORDER
lg.row_dimensions[4].height = 24
for r in range(5, 1006):
    for j in range(1, 16):
        c = lg.cell(r, j)
        c.font = font(9)
        if j in (10, 11, 12):
            c.number_format = PESO
        elif j in (2, 14):
            c.number_format = "yyyy-mm-dd hh:mm"
lg.freeze_panes = "B5"
lg.auto_filter.ref = "A4:O1005"
lg.conditional_formatting.add("A5:O1005", FormulaRule(formula=['$M5="Cancelled"'],
                              font=Font(name=FONT, color="9E9E9E", strike=True), fill=fill("FAFAFA")))
lg.conditional_formatting.add("M5:M1005", FormulaRule(formula=['$M5="Active"'],
                              font=Font(name=FONT, bold=True, color="1B5E20"), fill=fill("E8F5E9")))
lg.page_setup.orientation = "landscape"
lg.page_setup.fitToWidth = 1
lg.page_setup.fitToHeight = 0
lg.sheet_properties.pageSetUpPr.fitToPage = True

# =============================================================================
# MONTHLY SUMMARY (formulas only)
# =============================================================================
sm = wb.create_sheet("Monthly Summary")
sm.sheet_properties.tabColor = ACCENT
sm.sheet_view.showGridLines = False
banner(sm, "A1:I1", "A2:I2", "MONTHLY SUMMARY", "Students and total enrollment value per program, per month")
sm.row_dimensions[1].height = 30
heads = ["Month", "Toddlers\nStudents", "Toddlers\nTotal", "Preschool\nStudents", "Preschool\nTotal",
         "One-on-One\nStudents", "One-on-One\nTotal", "All\nStudents", "All Programs\nTotal"]
colors = ["37474F", "2E7D32", "2E7D32", "1E88E5", "1E88E5", "1A237E", "1A237E", "37474F", "37474F"]
for j, (h, col) in enumerate(zip(heads, colors)):
    sm.column_dimensions[L(j + 1)].width = 14 if j else 16
    c = sm.cell(4, j + 1, h)
    c.font = font(10, True, "FFFFFF"); c.fill = fill(col); c.alignment = CENTER; c.border = BORDER
sm.row_dimensions[4].height = 34
for i, m in enumerate(MONTHS):
    r = 5 + i
    vals = [m, f"={m}!G5", f"={m}!G4", f"={m}!O5", f"={m}!O4", f"={m}!V5", f"={m}!V4",
            f"=B{r}+D{r}+F{r}", f"=C{r}+E{r}+G{r}"]
    for j, v in enumerate(vals):
        c = sm.cell(r, j + 1, v)
        c.border = BORDER
        c.fill = fill("F5F7F9") if i % 2 else fill("FFFFFF")
        c.font = font(10, j in (0, 7, 8), "008000" if 1 <= j <= 6 else "000000")
        if j in (2, 4, 6, 8):
            c.number_format = PESO
            c.alignment = RIGHT
        elif j:
            c.alignment = Alignment(horizontal="center", vertical="center")
r = 17
sm.cell(r, 1, "TOTAL")
for j in range(2, 10):
    sm.cell(r, j, f"=SUM({L(j)}5:{L(j)}16)")
for j in range(1, 10):
    c = sm.cell(r, j)
    c.font = font(10, True, "FFFFFF"); c.fill = fill("37474F"); c.border = BORDER
    c.number_format = PESO if j in (3, 5, 7, 9) else "0"
    c.alignment = RIGHT if j in (3, 5, 7, 9) else Alignment(horizontal="center", vertical="center")
sm["A17"].alignment = LEFT
sm["A19"] = ("Note: a student enrolled for several months appears once in each covered month, "
             "so the 'Students' TOTAL row counts student-months.")
sm["A19"].font = font(9, False, GREY_TXT, True)
sm["A21"] = "Enrollment records (from Enrollment History)"
sm["A21"].font = font(10, True)
for k, (lbl, f) in enumerate([("Active", "=COUNTIF('Enrollment History'!$M:$M,\"Active\")"),
                              ("Cancelled", "=COUNTIF('Enrollment History'!$M:$M,\"Cancelled\")"),
                              ("Total submitted", "=COUNTA('Enrollment History'!$A$5:$A$1005)")]):
    sm.cell(22 + k, 1, lbl).font = font(10)
    c = sm.cell(22 + k, 2, f); c.font = font(10, True); c.alignment = Alignment(horizontal="center")

ch = BarChart()
ch.type = "col"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Enrollment Value per Month (₱)"
ch.y_axis.title = None
ch.height, ch.width = 8.5, 22
for col_idx in (3, 5, 7):
    data = Reference(sm, min_col=col_idx, min_row=4, max_row=16)
    ch.add_data(data, titles_from_data=True)
ch.set_categories(Reference(sm, min_col=1, min_row=5, max_row=16))
for s, colr in zip(ch.series, ["2E7D32", "1E88E5", "1A237E"]):
    s.graphicalProperties.solidFill = colr
    s.graphicalProperties.line.solidFill = colr
ch.roundedCorners = False
ch.legend.position = "b"
ch.legend.overlay = False
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.y_axis.numFmt = '"₱"#,##0'
sm.add_chart(ch, "K4")
sm.freeze_panes = "B5"

# =============================================================================
# STUDENT SEARCH (results written by VBA)
# =============================================================================
ss = wb.create_sheet("Student Search")
ss.sheet_properties.tabColor = "00897B"
ss.sheet_view.showGridLines = False
banner(ss, "A1:J1", "A2:J2", "STUDENT SEARCH", "Results across all 12 monthly sheets")
ss.row_dimensions[1].height = 30
ss["A3"] = "Search term:"
ss["A3"].font = font(10, True)
ss["B3"].font = font(10, True, "0000FF")
ss["D3"].font = font(10, False, GREY_TXT, True)
sh = ["Month", "Program Section", "No.", "Full Name", "Schedule", "Package", "Discount", "Price",
      "Enrollment ID", "Status"]
sw = [13, 20, 6, 26, 30, 11, 12, 13, 14, 11]
for j, (h, w) in enumerate(zip(sh, sw)):
    ss.column_dimensions[L(j + 1)].width = w
    c = ss.cell(5, j + 1, h)
    c.font = font(10, True, "FFFFFF"); c.fill = fill("00695C"); c.alignment = CENTER; c.border = BORDER
ss.row_dimensions[5].height = 22
for r in range(6, 506):
    for j in (7, 8):
        ss.cell(r, j).number_format = PESO
    for j in range(1, 11):
        ss.cell(r, j).font = font(9)
ss.freeze_panes = "A6"

# sheet order: Dashboard, months, Monthly Summary, Enrollment History, Student Search, Settings
order = ["Dashboard"] + MONTHS + ["Monthly Summary", "Enrollment History", "Student Search", "Settings"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = ws.title == "Dashboard"

wb.save(OUT)
print("saved", OUT)
