"""
Builds UpKidz_Enrollment_System.xlsx (v2) - the workbook half of the UpKidz
Enrollment Management System. The VBA half lives in UpKidzEnrollment.bas and
is embedded by build_xlsm.py.

Run:  python3 build_workbook.py [output_path]

Layout constants here MUST match the constants at the top of the .bas module
(FIRST_ROW, LAST_ROW, section start columns, history columns, defined names).
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, LineChart, Reference

OUT = sys.argv[1] if len(sys.argv) > 1 else "UpKidz_Enrollment_System.xlsx"

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

FIRST_ROW, LAST_ROW = 11, 110         # 100 enrollment rows per section
LOG_FIRST, LOG_LAST = 5, 1005
FONT = "Arial"
PESO = '"₱"#,##0.00;-"₱"#,##0.00;"-"'
PESO0 = '"₱"#,##0;-"₱"#,##0;"-"'
PCT = '0.0%'

# ---- client palette -------------------------------------------------------
TOD_DARK, TOD_LIGHT = "375623", "E2EFDA"
PRE_DARK, PRE_LIGHT = "1F4E78", "DDEBF7"
TUT_DARK, TUT_LIGHT = "203764", "D9E1F2"
ACCENT = "FFC107"
NAVY = TUT_DARK
WHITE = "FFFFFF"
CALC_FILL = "F2F2F2"
GREY_TXT = "595959"
INPUT_TXT = "1F4E78"

SECTIONS = [
    dict(key="Toddler", title="TODDLERS GROUP", caption="Toddler PlayTrack", start=2,
         cols=["No.", "Full Name", "Package", "Discount", "Price", "Per Month"],
         dark=TOD_DARK, light=TOD_LIGHT),
    dict(key="Preschool", title="PRESCHOOL GROUP", caption="Preschool PlayTrack", start=10,
         cols=["No.", "Full Name", "Package", "Discount", "Price", "Per Month"],
         dark=PRE_DARK, light=PRE_LIGHT),
    dict(key="Tutor", title="ONE-ON-ONE TUTORING", caption="One-on-One Tutoring", start=18,
         cols=["No.", "Full Name", "Package", "Price", "Per Month"],
         dark=TUT_DARK, light=TUT_LIGHT),
]
WIDTHS = {"No.": 5.5, "Full Name": 27, "Package": 16, "Discount": 11.5, "Price": 13, "Per Month": 13}
SECTION_NAMES = ["Toddlers Group", "Preschool Group", "One-on-One Tutoring"]

# client's price list (Settings)
PROGRAMS = [("Toddler PlayTrack", "Toddlers Group"), ("Toddler PlayTrack VIP", "Toddlers Group"),
            ("Preschool PlayTrack", "Preschool Group"), ("One-on-One Tutoring 4x", "One-on-One Tutoring"),
            ("One-on-One Tutoring 2x", "One-on-One Tutoring")]
PACKAGES = [
    ("Toddler PlayTrack", "3 Months", 3, 10000), ("Toddler PlayTrack", "6 Months", 6, 18000),
    ("Toddler PlayTrack VIP", "3 Months", 3, 15000), ("Toddler PlayTrack VIP", "6 Months", 6, 24000),
    ("Preschool PlayTrack", "3 Months", 3, 15000), ("Preschool PlayTrack", "6 Months", 6, 24000),
    ("One-on-One Tutoring 4x", "1 Month", 1, 3500), ("One-on-One Tutoring 4x", "2 Months", 2, 6500),
    ("One-on-One Tutoring 2x", "1 Month", 1, 2000), ("One-on-One Tutoring 2x", "3 Months", 3, 5000),
    ("One-on-One Tutoring 2x", "10 Months", 10, 15700),
]

thin = Side(style="thin", color="BFBFBF")
med = Side(style="medium", color=NAVY)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def font(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", indent=1)
RIGHT = Alignment(horizontal="right", vertical="center")
MID = Alignment(horizontal="center", vertical="center")

wb = Workbook()
dn = wb.defined_names


def name(n, ref):
    dn[n] = DefinedName(n, attr_text=ref)


def banner(ws, rng1, rng2, title, sub):
    ws.merge_cells(rng1)
    ws.merge_cells(rng2)
    for row in ws[rng1] + ws[rng2]:
        for c in row:
            c.fill = fill(NAVY)
    a = ws[rng1.split(":")[0]]
    a.value, a.font, a.alignment = title, font(18, True, ACCENT), MID
    b = ws[rng2.split(":")[0]]
    b.value, b.font, b.alignment = sub, font(10, True, WHITE), MID
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20


def header_cell(c, text, color=NAVY, size=10):
    c.value = text
    c.font = font(size, True, WHITE)
    c.fill = fill(color)
    c.alignment = CENTER
    c.border = BORDER


# =============================================================================
# SETTINGS
# =============================================================================
st = wb.active
st.title = "Settings"
st.sheet_properties.tabColor = "808080"
st.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHI", [26, 14, 10, 15, 36, 3, 26, 22, 3]):
    st.column_dimensions[col].width = w
st.merge_cells("A1:H1")
st["A1"] = "SYSTEM SETTINGS & PRICE LIST"
st["A1"].font = font(14, True, ACCENT)
st["A1"].fill = fill(NAVY)
st["A1"].alignment = LEFT
st.row_dimensions[1].height = 28
cfg = [(2, "School Year", 2026, "0", "Shown in monthly sheet titles."),
       (3, "Royalty rate", 0.05, "0%", "Royalty = Equivalent Sales × rate (5%)."),
       (4, "Marketing contribution rate", 0.02, "0%", "Marketing = Equivalent Sales × rate (2%).")]
for r, k, v, fmt, note in cfg:
    st.cell(r, 1, k).font = font(10, True)
    c = st.cell(r, 2, v)
    c.font, c.border, c.alignment, c.number_format = font(10, True, INPUT_TXT), BORDER, MID, fmt
    st.cell(r, 3, note).font = font(9, color=GREY_TXT, italic=True)

PKG_HDR, PKG_FIRST, PKG_LAST = 6, 7, 40
for j, h in enumerate(["Program", "Package", "Months", "Program Fee", "Key (auto)"]):
    header_cell(st.cell(PKG_HDR, j + 1), h)
for r in range(PKG_FIRST, PKG_LAST + 1):
    i = r - PKG_FIRST
    if i < len(PACKAGES):
        for j, v in enumerate(PACKAGES[i]):
            st.cell(r, j + 1, v)
    for j in range(1, 5):
        c = st.cell(r, j)
        c.font, c.border = font(10, color=INPUT_TXT), BORDER
    st.cell(r, 3).alignment = MID
    st.cell(r, 4).number_format = PESO0
    k = st.cell(r, 5, f'=IF(A{r}="","",TRIM(A{r})&"|"&TRIM(B{r}))')
    k.font, k.border, k.fill = font(9, color=GREY_TXT), BORDER, fill(CALC_FILL)
st.cell(PKG_LAST + 2, 1, "Keep each program's packages in consecutive rows. Program names must match the "
        "Programs table. Months = how many monthly sheets the package covers (1–12).").font = font(9, color=GREY_TXT, italic=True)

# months G6:G18
header_cell(st["G6"], "Months")
for i, m in enumerate(MONTHS):
    c = st.cell(7 + i, 7, m)
    c.border, c.font = BORDER, font(10)
# programs table G21:H31
PROG_HDR, PROG_FIRST, PROG_LAST = 21, 22, 31
header_cell(st.cell(PROG_HDR, 7), "Program")
header_cell(st.cell(PROG_HDR, 8), "Monthly-sheet Section")
for r in range(PROG_FIRST, PROG_LAST + 1):
    i = r - PROG_FIRST
    if i < len(PROGRAMS):
        st.cell(r, 7, PROGRAMS[i][0])
        st.cell(r, 8, PROGRAMS[i][1])
    for cc in (7, 8):
        c = st.cell(r, cc)
        c.border, c.font = BORDER, font(10, color=INPUT_TXT)
st.cell(PROG_LAST + 1, 7, "Section decides which table the student goes to on the monthly sheets.").font = font(9, color=GREY_TXT, italic=True)
dv = DataValidation(type="list", formula1='"' + ",".join(SECTION_NAMES) + '"', allow_blank=True)
st.add_data_validation(dv)
dv.add(f"H{PROG_FIRST}:H{PROG_LAST}")
dv = DataValidation(type="list", formula1="ProgramList", allow_blank=True,
                    error="Program must exist in the Programs table (G22:G31).", errorTitle="Unknown program")
dv.showErrorMessage = True
st.add_data_validation(dv)
dv.add(f"A{PKG_FIRST}:A{PKG_LAST}")
st.freeze_panes = "A7"

name("MonthList", "Settings!$G$7:$G$18")
name("ProgTable", f"Settings!$G${PROG_FIRST}:$H${PROG_LAST}")
name("ProgramList", f"OFFSET(Settings!$G${PROG_FIRST},0,0,MAX(1,COUNTA(Settings!$G${PROG_FIRST}:$G${PROG_LAST})),1)")
name("PkgProgram", f"Settings!$A${PKG_FIRST}:$A${PKG_LAST}")
name("PkgName", f"Settings!$B${PKG_FIRST}:$B${PKG_LAST}")
name("PkgMonths", f"Settings!$C${PKG_FIRST}:$C${PKG_LAST}")
name("PkgFee", f"Settings!$D${PKG_FIRST}:$D${PKG_LAST}")
name("PkgKey", f"Settings!$E${PKG_FIRST}:$E${PKG_LAST}")
name("cfgSchoolYear", "Settings!$B$2")
name("cfgRoyalty", "Settings!$B$3")
name("cfgMarketing", "Settings!$B$4")
name("PackageChoices",
     f"OFFSET(Settings!$B${PKG_FIRST},MATCH(TRIM(Dashboard!$C$6),Settings!$A${PKG_FIRST}:$A${PKG_LAST},0)-1,0,"
     f"MAX(1,COUNTIF(Settings!$A${PKG_FIRST}:$A${PKG_LAST},TRIM(Dashboard!$C$6))),1)")
name("selMonths", f"IFERROR(INDEX(Settings!$C${PKG_FIRST}:$C${PKG_LAST},MATCH(TRIM(Dashboard!$C$6)&\"|\"&TRIM(Dashboard!$C$9),"
     f"Settings!$E${PKG_FIRST}:$E${PKG_LAST},0)),0)")
name("HistSID", f"'Enrollment History'!$B${LOG_FIRST}:$B${LOG_LAST}")
name("HistName", f"'Enrollment History'!$G${LOG_FIRST}:$G${LOG_LAST}")

# =============================================================================
# DASHBOARD
# =============================================================================
ds = wb.create_sheet("Dashboard", 0)
ds.sheet_properties.tabColor = ACCENT
ds.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJK", [2, 22, 22, 22, 36, 3, 20, 16, 16, 16, 2]):
    ds.column_dimensions[col].width = w
banner(ds, "B1:J1", "B2:J2", "UPKIDZ PHILIPPINES", "Enrollment Management System  •  Dashboard")
ds.row_dimensions[3].height = 8


def card_header(ws, rng, text):
    ws.merge_cells(rng)
    for row in ws[rng]:
        for x in row:
            x.fill = fill(NAVY)
    c = ws[rng.split(":")[0]]
    c.value, c.font, c.alignment = text, font(11, True, ACCENT), LEFT


card_header(ds, "B4:E4", "ENROLLMENT FORM")
ds.row_dimensions[4].height = 24
ds["B5"] = "Mode"
ds["B5"].font = font(9, True, GREY_TXT)
ds["C5"].font = font(10, True, "C00000")       # editing enrollment ID (set by VBA)
ds.merge_cells("D5:E5")
ds["D5"] = ('=IF(C5="","NEW ENROLLMENT — white cells are for typing, grey cells are automatic",'
            '"EDITING "&C5&" — change fields, then click UPDATE ENROLLMENT")')
ds["D5"].font = font(9, True, GREY_TXT, True)

fields = [  # row, label, name, kind
    (6, "Program", "inProgram", "input"),
    (7, "Full Name", "inName", "input"),
    (8, "Student", "outStudent", "calc"),
    (9, "Package", "inPackage", "input"),
    (10, "Start Month", "inStartMonth", "input"),
    (11, "Duration", "outDuration", "calc"),
    (12, "Program Fee (₱)", "outFee", "calc"),
    (13, "Discount (₱)", "inDiscount", "input"),
    (14, "Net Price (₱)", "outNetPrice", "calc"),
    (15, "Enrollment Status", "inStatus", "input"),
    (16, "Last Month Attended", "inDropMonth", "input"),
    (17, "Status Check", "outStatus", "calc"),
]
for r, label, nm, kind in fields:
    ds.row_dimensions[r].height = 22
    lc = ds.cell(r, 2, label)
    lc.font, lc.alignment, lc.fill, lc.border = font(10, True, NAVY), LEFT, fill(TUT_LIGHT), BORDER
    ds.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    for cc in (3, 4):
        x = ds.cell(r, cc)
        x.border = BORDER
        x.fill = fill(WHITE if kind == "input" else CALC_FILL)
    v = ds.cell(r, 3)
    v.font = font(11, kind == "calc", INPUT_TXT if kind == "input" else "000000")
    v.alignment = LEFT
    name(nm, f"Dashboard!$C${r}")
name("inEditID", "Dashboard!$C$5")
# input cells get an amber left edge so they stand out
for r, _, _, kind in fields:
    if kind == "input":
        ds.cell(r, 3).border = Border(left=Side(style="thick", color=ACCENT), right=thin, top=thin, bottom=thin)

ds["C8"] = ('=IF(TRIM(C7)="","",IFERROR("Returning student · "&IFERROR(INDEX(HistSID,MATCH(TRIM(C7),HistName,0)),'
            'INDEX(ArchSID,MATCH(TRIM(C7),ArchName,0)))&" (will be recorded as re-enrollment)","New student"))')
ds["C11"] = ('=IF(OR(C10="",selMonths=0),"",IF(ISNA(MATCH(C10,MonthList,0)),"Invalid start month",'
             'C10&IF(selMonths>1,"–"&INDEX(MonthList,MOD(MATCH(C10,MonthList,0)+selMonths-2,12)+1),"")'
             '&" ("&selMonths&IF(selMonths=1," month)"," months)")))')
ds["C12"] = '=IF(selMonths=0,"",INDEX(PkgFee,MATCH(TRIM(C6)&"|"&TRIM(C9),PkgKey,0)))'
ds["C14"] = '=IF(C12="","",C12-N(C13))'
ds["C17"] = ('=IF(C6="","⚠ Select a program",IF(TRIM(C7)="","⚠ Enter the full name",'
             'IF(C9="","⚠ Select a package",IF(selMonths=0,"⚠ Package not found in Settings",'
             'IF(C10="","⚠ Select the start month",IF(N(C13)>N(C12),"⚠ Discount exceeds the program fee",'
             'IF(AND(C15="Dropped",C5=""),"⚠ Dropped is only for UPDATE (load an enrollment first)",'
             'IF(AND(C15="Dropped",C16=""),"⚠ Select the last month attended",'
             'IF(C5="","✔ Ready to submit","✔ Ready to update")))))))))')
for a in ("C12", "C13", "C14"):
    ds[a].number_format = PESO
ds["C14"].font = font(12, True, TOD_DARK)

hints = {
    6: "Choose from the Settings program list",
    7: "Student's complete name",
    8: "Automatic — finds the student in Enrollment History",
    9: '=IF(selMonths>0,"Package covers "&selMonths&" month(s)",IF(C6="","Choose a program first","Choose a package"))',
    10: "First month the student attends",
    11: ('=IF(AND(selMonths>0,ISNUMBER(MATCH(C10,MonthList,0))),IF(MATCH(C10,MonthList,0)+selMonths-1>12,'
         '"⚠ "&(MATCH(C10,MonthList,0)+selMonths-13)&" month(s) go to next year (End of Year sheet)",'
         '"Automatic — from the package"),"Automatic — from the package")'),
    12: "Automatic — from the Settings price list",
    13: '=IF(N(C13)>N(C12),"⚠ Discount is greater than the fee!","Leave blank if none")',
    14: "Automatic — Program Fee − Discount",
    15: "Active (blank) or Dropped — Dropped is used with UPDATE",
    16: "Last month the student attends; removed from the next month on",
    17: "Must show ✔ before clicking the button",
}
for r, h in hints.items():
    c = ds.cell(r, 5, h)
    c.font = font(9, False, GREY_TXT, True)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)

ds.conditional_formatting.add("C17", FormulaRule(formula=['LEFT($C$17,1)="✔"'], fill=fill(TOD_LIGHT),
                              font=Font(name=FONT, bold=True, color=TOD_DARK)))
ds.conditional_formatting.add("C17", FormulaRule(formula=['LEFT($C$17,1)="⚠"'], fill=fill("FFF2CC"),
                              font=Font(name=FONT, bold=True, color="C00000")))
ds.conditional_formatting.add("C8", FormulaRule(formula=['LEFT($C$8,9)="Returning"'], fill=fill(PRE_LIGHT),
                              font=Font(name=FONT, bold=True, color=PRE_DARK)))
ds.conditional_formatting.add("E9:E13", FormulaRule(formula=['LEFT(E9,1)="⚠"'],
                              font=Font(name=FONT, bold=True, color="C00000")))
ds.conditional_formatting.add("C15:C16", FormulaRule(formula=['$C$15="Dropped"'], fill=fill("FFF2CC"),
                              font=Font(name=FONT, bold=True, color="C00000")))


def dv_add(ws, cell, **kw):
    show = kw.pop("show", True)
    d = DataValidation(**kw)
    d.showErrorMessage = show
    ws.add_data_validation(d)
    d.add(cell)


dv_add(ds, "C6", type="list", formula1="ProgramList", allow_blank=True,
       error="Choose a program from the list (edit the list on the Settings sheet).", errorTitle="Invalid program")
dv_add(ds, "C9", type="list", formula1="PackageChoices", allow_blank=True,
       error="Choose a package listed for the selected program (see Settings).", errorTitle="Invalid package")
dv_add(ds, "C10", type="list", formula1="MonthList", allow_blank=True,
       error="Choose a month from January to December.", errorTitle="Invalid month")
dv_add(ds, "C13", type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
       error="Discount must be zero or more.", errorTitle="Invalid discount")
dv_add(ds, "C15", type="list", formula1='"Active,Dropped"', allow_blank=True,
       error="Choose Active or Dropped.", errorTitle="Invalid status")
dv_add(ds, "C16", type="list", formula1="MonthList", allow_blank=True,
       error="Choose a month from January to December.", errorTitle="Invalid month")

for rr in (18, 19, 20, 21):
    ds.row_dimensions[rr].height = 22
ds.merge_cells("B19:E20")
ds["B19"] = "Buttons appear here after running the macro  InstallUpKidzSystem  (see README)."
ds["B19"].font = font(9, False, GREY_TXT, True)
ds["B19"].alignment = CENTER

# ---- right column ------------------------------------------------------------
card_header(ds, "G4:J4", "MANAGE ENROLLMENTS")


def labeled_input(r, label, nm):
    ds.row_dimensions[r].height = 22
    c = ds.cell(r, 7, label)
    c.font, c.fill, c.border, c.alignment = font(10, True, NAVY), fill(TUT_LIGHT), BORDER, LEFT
    ds.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
    for cc in (8, 9, 10):
        x = ds.cell(r, cc)
        x.fill, x.border = fill(WHITE), BORDER
    ds.cell(r, 8).border = Border(left=Side(style="thick", color=ACCENT), right=thin, top=thin, bottom=thin)
    ds.cell(r, 8).font, ds.cell(r, 8).alignment = font(11, False, INPUT_TXT), LEFT
    name(nm, f"Dashboard!$H${r}")


labeled_input(6, "Enrollment ID", "inManageID")
ds.merge_cells("G7:J7")
ds["G7"] = "Edit / drop: type the ID → LOAD FOR EDIT → change fields → UPDATE ENROLLMENT."
ds["G7"].font = font(8, False, GREY_TXT, True)
ds.row_dimensions[8].height = 30
labeled_input(10, "Student", "inSearch")
ds.merge_cells("G11:J11")
ds["G11"] = "Name, Student ID or Enrollment ID. PROFILE shows history; RE-ENROLL fills the form."
ds["G11"].font = font(8, False, GREY_TXT, True)
ds.row_dimensions[12].height = 30

card_header(ds, "G14:J14", "MONTHLY SNAPSHOT")
ds["G15"] = "Month"
ds["G15"].font, ds["G15"].fill, ds["G15"].border, ds["G15"].alignment = font(10, True, NAVY), fill(TUT_LIGHT), BORDER, LEFT
ds.merge_cells("H15:J15")
ds["H15"] = "January"
for cc in (8, 9, 10):
    ds.cell(15, cc).fill, ds.cell(15, cc).border = fill(WHITE), BORDER
ds["H15"].border = Border(left=Side(style="thick", color=ACCENT), right=thin, top=thin, bottom=thin)
ds["H15"].font, ds["H15"].alignment = font(11, True, INPUT_TXT), LEFT
dv_add(ds, "H15", type="list", formula1="MonthList", allow_blank=False, show=False)
for j, h in enumerate(["Program", "Students", "Package Value", "Per Month"]):
    header_cell(ds.cell(16, 7 + j), h, color="404040", size=9)
SNAP = [("Toddlers Group", SECTIONS[0], "C", "F8", "G8"),
        ("Preschool Group", SECTIONS[1], "K", "N8", "O8"),
        ("One-on-One", SECTIONS[2], "S", "U8", "V8")]
IND = "INDIRECT(\"'\"&$H$15&\"'!{}\")"
for i, (lbl, sec, namecol, pv, pm) in enumerate(SNAP):
    r = 17 + i
    c = ds.cell(r, 7, lbl)
    c.font, c.fill, c.border, c.alignment = font(10, True, WHITE), fill(sec["dark"]), BORDER, LEFT
    c = ds.cell(r, 8, "=IFERROR(COUNTA(" + IND.format(f"{namecol}{FIRST_ROW}:{namecol}{LAST_ROW}") + "),0)")
    c.font, c.alignment, c.border = font(11, True), MID, BORDER
    c = ds.cell(r, 9, "=IFERROR(" + IND.format(pv) + ",0)")
    c.font, c.alignment, c.border, c.number_format = font(10), RIGHT, BORDER, PESO
    c = ds.cell(r, 10, "=IFERROR(" + IND.format(pm) + ",0)")
    c.font, c.alignment, c.border, c.number_format = font(10, True), RIGHT, BORDER, PESO
c = ds.cell(20, 7, "TOTAL")
for cc, f in ((8, "=SUM(H17:H19)"), (9, "=SUM(I17:I19)"), (10, "=SUM(J17:J19)")):
    x = ds.cell(20, cc, f)
    x.number_format = PESO if cc > 8 else "0"
    x.alignment = RIGHT if cc > 8 else MID
for cc in range(7, 11):
    x = ds.cell(20, cc)
    x.font, x.fill, x.border = font(10, True, NAVY), fill(TUT_LIGHT), BORDER
ds["G20"].alignment = LEFT
KPIS = [  # row, label, formula, fmt
    (22, "Total Students", "=IFERROR(" + IND.format("B5") + ",0)", "0"),
    (23, "Equivalent Sales", "=IFERROR(" + IND.format("D5") + ",0)", PESO),
    (24, '="Royalty ("&TEXT(cfgRoyalty,"0%")&")"', "=IFERROR(" + IND.format("F5") + ",0)", PESO),
    (25, '="Marketing Contribution ("&TEXT(cfgMarketing,"0%")&")"', "=IFERROR(" + IND.format("I5") + ",0)", PESO),
    (26, "TOTAL DUES (Royalty + Marketing)", "=IFERROR(" + IND.format("L5") + ",0)", PESO),
    (27, "Retention Rate", "=IFERROR(" + IND.format("O5") + ",\"N/A\")", PCT),
    (28, "Dropped (last month attended)", "=IFERROR(" + IND.format("T5") + ",0)", "0"),
]
for r, lbl, f, fmt in KPIS:
    ds.row_dimensions[r].height = 21
    ds.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
    ds.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    c = ds.cell(r, 7, lbl)
    c.font, c.fill, c.border, c.alignment = font(10, True, WHITE), fill(NAVY), BORDER, LEFT
    ds.cell(r, 8).border = BORDER
    c = ds.cell(r, 9, f)
    c.font, c.number_format, c.alignment, c.border = font(12, True, NAVY), fmt, RIGHT, BORDER
    ds.cell(r, 10).border = BORDER
ds.merge_cells("G29:J29")
ds["G29"] = "=IFERROR(" + IND.format("O6") + ",\"\")"
ds["G29"].font = font(8, False, GREY_TXT, True)
ds["G29"].alignment = Alignment(horizontal="right", vertical="center")

ds["B22"] = "HOW TO USE"
ds["B22"].font = font(10, True, NAVY)
steps = [
    "1. Type in the white cells (amber edge). Grey cells — fee, duration, net price, per month — fill in automatically.",
    "2. When Status Check shows ✔, click SUBMIT ENROLLMENT. The student is added to every month the package covers.",
    "3. Re-enroll: type the name in Student → RE-ENROLL → choose the new program/package → SUBMIT. History links both.",
    "4. Drop: Enrollment ID → LOAD FOR EDIT → Status = Dropped + Last Month Attended → UPDATE (removed from the next month on).",
    "5. Cancel (entered by mistake): Enrollment ID → CANCEL ENROLLMENT. PROFILE shows a student's full history.",
]
for i, s in enumerate(steps):
    r = 23 + i
    ds.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    c = ds.cell(r, 2, s)
    c.font = font(9, False, "262626")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    ds.row_dimensions[r].height = 24
ds.page_setup.orientation = "landscape"
ds.page_setup.fitToWidth = 1
ds.page_setup.fitToHeight = 1
ds.sheet_properties.pageSetUpPr.fitToPage = True
ds.print_area = "A1:K30"


# =============================================================================
# MONTHLY SHEETS
# =============================================================================
KPI_BLOCKS = [  # (first col, last col, label, value formula key, fmt)
    ("B", "C", "TOTAL STUDENTS", "students", "0"),
    ("D", "E", "EQUIVALENT SALES", "equiv", PESO),
    ("F", "H", "ROYALTY", "royalty", PESO),
    ("I", "K", "MARKETING CONTRIBUTION", "marketing", PESO),
    ("L", "N", "TOTAL DUES", "dues", PESO),
    ("O", "S", "RETENTION RATE", "retention", PCT),
    ("T", "V", "DROPPED", "dropped", "0"),
]


def build_month(ws, idx, month):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.column_dimensions["A"].width = 1.5
    prev = MONTHS[idx - 1] if idx > 0 else None
    banner(ws, "B1:V1", "B2:V2", "UPKIDZ PHILIPPINES", "")
    ws["B2"] = f'="{month.upper()} "&cfgSchoolYear&"  •  MONTHLY ENROLLMENT RECORD"'
    ws.row_dimensions[3].height = 6
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 30
    ws.row_dimensions[6].height = 15
    ws.row_dimensions[7].height = 8
    ws.row_dimensions[8].height = 26
    ws.row_dimensions[9].height = 16
    ws.row_dimensions[10].height = 22

    # section columns
    for sec in SECTIONS:
        s, n = sec["start"], len(sec["cols"])
        for j, h in enumerate(sec["cols"]):
            ws.column_dimensions[L(s + j)].width = WIDTHS[h]
        ws.column_dimensions[L(s + n)].width = 12
        ws.column_dimensions[L(s + n)].hidden = True       # Enrollment ID
        if sec is not SECTIONS[-1]:
            ws.column_dimensions[L(s + n + 1)].width = 2.5  # spacer
    ws.column_dimensions["Z"].width = 10
    ws.column_dimensions["Z"].hidden = True                # helper cells

    # helpers: Z4 retained, Z5 students last month, Z6 dropped this month
    nm = [f"{c}{FIRST_ROW}:{c}{LAST_ROW}" for c in ("C", "K", "S")]
    if prev:
        parts = []
        for pc in nm:
            pr = f"{prev}!{pc}"
            parts.append(f"SUMPRODUCT(({pr}<>\"\")*((COUNTIF($C${FIRST_ROW}:$C${LAST_ROW},{pr})"
                         f"+COUNTIF($K${FIRST_ROW}:$K${LAST_ROW},{pr})+COUNTIF($S${FIRST_ROW}:$S${LAST_ROW},{pr}))>0))")
        ws["Z4"] = "=" + "+".join(parts)
        ws["Z5"] = f"={prev}!B5"
    else:
        ws["Z4"] = 0
        ws["Z5"] = 0
    ws["Z6"] = (f"=COUNTIFS('Enrollment History'!$R${LOG_FIRST}:$R${LOG_LAST},\"Dropped\","
                f"'Enrollment History'!$S${LOG_FIRST}:$S${LOG_LAST},\"{month}\","
                f"'Enrollment History'!$T${LOG_FIRST}:$T${LOG_LAST},cfgSchoolYear)")

    vals = {
        "students": f"=COUNTA(C{FIRST_ROW}:C{LAST_ROW})+COUNTA(K{FIRST_ROW}:K{LAST_ROW})+COUNTA(S{FIRST_ROW}:S{LAST_ROW})",
        "equiv": "=G8+O8+V8",
        "royalty": "=D5*cfgRoyalty",
        "marketing": "=D5*cfgMarketing",
        "dues": "=F5+I5",
        "retention": '=IF(N(Z5)=0,"N/A",Z4/Z5)',
        "dropped": "=Z6",
    }
    captions = {
        "students": "All programs this month",
        "equiv": "Σ Price ÷ package months",
        "royalty": '=TEXT(cfgRoyalty,"0%")&" of equivalent sales"',
        "marketing": '=TEXT(cfgMarketing,"0%")&" of equivalent sales"',
        "dues": "Royalty + Marketing Contribution",
        "retention": (f'=IF(N(Z5)=0,"No students last month",Z4&" of "&Z5&" students from {prev} stayed")'
                      if prev else '="First month — no previous month"'),
        "dropped": "Last month attended = this month",
    }
    for c1, c2, label, key, fmt in KPI_BLOCKS:
        for r in (4, 5, 6):
            ws.merge_cells(f"{c1}{r}:{c2}{r}")
        lab = ws[f"{c1}4"]
        lab.value, lab.font, lab.alignment = label, font(9, True, ACCENT), MID
        v = ws[f"{c1}5"]
        v.value, v.font, v.alignment, v.number_format = vals[key], font(16, True, NAVY), MID, fmt
        cap = ws[f"{c1}6"]
        cap.value, cap.font, cap.alignment = captions[key], font(8, False, GREY_TXT, True), MID
        first, last = ws[f"{c1}4"].column, ws[f"{c2}4"].column
        for cc in range(first, last + 1):
            ws.cell(4, cc).fill = fill(NAVY)
            ws.cell(5, cc).fill = fill(WHITE)
            ws.cell(6, cc).fill = fill(TUT_LIGHT)
            ws.cell(5, cc).border = Border(left=Side(style="medium", color=ACCENT) if cc == first else None,
                                           right=Side(style="medium", color=ACCENT) if cc == last else None)
            ws.cell(6, cc).border = Border(bottom=Side(style="medium", color=ACCENT),
                                           left=Side(style="medium", color=ACCENT) if cc == first else None,
                                           right=Side(style="medium", color=ACCENT) if cc == last else None)
            ws.cell(4, cc).border = Border(top=Side(style="medium", color=ACCENT),
                                           left=Side(style="medium", color=ACCENT) if cc == first else None,
                                           right=Side(style="medium", color=ACCENT) if cc == last else None)

    for sec in SECTIONS:
        s, n = sec["start"], len(sec["cols"])
        e = s + n - 1                    # Per Month column
        price_c, pm_c, name_c = L(e - 1), L(e), L(s + 1)
        # row 8: title + totals (top-right corner)
        ws.merge_cells(start_row=8, start_column=s, end_row=8, end_column=e - 3)
        t = ws.cell(8, s, sec["title"])
        t.font, t.alignment = font(12, True, WHITE), LEFT
        lab = ws.cell(8, e - 2, "TOTALS")
        lab.font, lab.alignment = font(9, True, ACCENT), RIGHT
        for col_letter, cc in ((price_c, e - 1), (pm_c, e)):
            tot = ws.cell(8, cc, f"=SUM({col_letter}{FIRST_ROW}:{col_letter}{LAST_ROW})")
            tot.font, tot.fill, tot.number_format, tot.alignment = font(10, True, sec["dark"]), fill(WHITE), PESO, RIGHT
            tot.border = Border(left=Side(style="medium", color=ACCENT), right=Side(style="medium", color=ACCENT),
                                top=Side(style="medium", color=ACCENT), bottom=Side(style="medium", color=ACCENT))
        for cc in range(s, e - 1):
            ws.cell(8, cc).fill = fill(sec["dark"])
        # row 9: caption + student count, captions under totals
        ws.merge_cells(start_row=9, start_column=s, end_row=9, end_column=e - 3)
        p = ws.cell(9, s, f'="{sec["caption"]}  •  Students: "&COUNTA({name_c}{FIRST_ROW}:{name_c}{LAST_ROW})')
        p.font, p.alignment = font(9, True, sec["dark"]), LEFT
        ws.cell(9, e - 1, "Package value").font = font(8, True, sec["dark"])
        ws.cell(9, e, "Per month").font = font(8, True, sec["dark"])
        ws.cell(9, e - 1).alignment = ws.cell(9, e).alignment = RIGHT
        for cc in range(s, e + 1):
            ws.cell(9, cc).fill = fill(sec["light"])
        # row 10: headers
        for j, h in enumerate(sec["cols"]):
            c = ws.cell(10, s + j, h)
            c.font, c.fill, c.alignment = font(10, True, WHITE), fill(sec["dark"]), CENTER
            c.border = Border(left=Side(style="thin", color=WHITE), right=Side(style="thin", color=WHITE),
                              bottom=Side(style="medium", color=ACCENT))
        ws.cell(10, s + n, "Enrollment ID").font = font(8, True, GREY_TXT)
        # data rows
        for r in range(FIRST_ROW, LAST_ROW + 1):
            band = fill(sec["light"]) if (r - FIRST_ROW) % 2 else fill(WHITE)
            for j, h in enumerate(sec["cols"]):
                c = ws.cell(r, s + j)
                c.fill, c.border, c.font = band, BORDER, font(10)
                if h == "No.":
                    c.value = f'=IF({name_c}{r}="","",COUNTA({name_c}${FIRST_ROW}:{name_c}{r}))'
                    c.alignment, c.font = MID, font(10, True, sec["dark"])
                elif h in ("Discount", "Price", "Per Month"):
                    c.number_format, c.alignment = PESO, RIGHT
                    if h == "Per Month":
                        c.font = font(10, True, sec["dark"])
                elif h == "Package":
                    c.alignment = MID
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.cell(r, s + n).font = font(8, color=GREY_TXT)
            ws.row_dimensions[r].height = 17
        for cc in range(s, e + 1):
            ws.cell(LAST_ROW, cc).border = Border(left=thin, right=thin, top=thin,
                                                  bottom=Side(style="medium", color=sec["dark"]))
    foot = ws.cell(LAST_ROW + 2, 2, "Records are added automatically from the Dashboard. Use the Dashboard "
                   "(Update / Cancel / Re-enroll) so every month stays consistent. Retention counts students "
                   "(by name) who were also enrolled last month.")
    foot.font = font(8, False, GREY_TXT, True)
    ws.freeze_panes = f"B{FIRST_ROW}"
    ws.print_title_rows = "1:10"
    ws.print_area = f"A1:V{LAST_ROW}"
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
    ws.sheet_properties.tabColor = [TOD_DARK, PRE_DARK, TUT_DARK][i % 3]
    build_month(ws, i, m)

# =============================================================================
# ENROLLMENT HISTORY (student profiles; written by VBA)
# =============================================================================
lg = wb.create_sheet("Enrollment History")
lg.sheet_properties.tabColor = ACCENT
lg.sheet_view.showGridLines = False
HIST = [("Enrollment ID", 13), ("Student ID", 11), ("Date Submitted", 16), ("Type", 13), ("Program", 22),
        ("Section", 18), ("Full Name", 26), ("Package", 11), ("Months", 8), ("Start Month", 11),
        ("Start Year", 9), ("Months Covered (this year)", 34), ("Next-Year Months", 26), ("Program Fee", 12),
        ("Discount", 11), ("Net Price", 12), ("Per Month", 12), ("Status", 10), ("Last Month Attended", 12),
        ("Last Month Year", 9), ("Last Updated", 16), ("Remarks", 40)]
banner(lg, "A1:V1", "A2:V2", "ENROLLMENT HISTORY",
       "Every enrollment, re-enrollment, drop and cancellation — one row per enrollment, linked by Student ID")
for j, (h, w) in enumerate(HIST):
    lg.column_dimensions[L(j + 1)].width = w
    header_cell(lg.cell(4, j + 1), h)
lg.row_dimensions[4].height = 24
for r in range(LOG_FIRST, LOG_LAST + 1):
    for j in range(1, len(HIST) + 1):
        c = lg.cell(r, j)
        c.font = font(9)
        if j in (14, 15, 16, 17):
            c.number_format = PESO
        elif j in (3, 21):
            c.number_format = "yyyy-mm-dd hh:mm"
lg.freeze_panes = "C5"
lg.auto_filter.ref = f"A4:V{LOG_LAST}"
lg.conditional_formatting.add(f"A{LOG_FIRST}:V{LOG_LAST}", FormulaRule(
    formula=[f'$R{LOG_FIRST}="Cancelled"'], font=Font(name=FONT, color="A6A6A6", strike=True)))
lg.conditional_formatting.add(f"R{LOG_FIRST}:T{LOG_LAST}", FormulaRule(
    formula=[f'$R{LOG_FIRST}="Dropped"'], font=Font(name=FONT, bold=True, color="C00000"), fill=fill("FFF2CC")))
lg.conditional_formatting.add(f"R{LOG_FIRST}:R{LOG_LAST}", FormulaRule(
    formula=[f'$R{LOG_FIRST}="Active"'], font=Font(name=FONT, bold=True, color=TOD_DARK), fill=fill(TOD_LIGHT)))
lg.conditional_formatting.add(f"D{LOG_FIRST}:D{LOG_LAST}", FormulaRule(
    formula=[f'$D{LOG_FIRST}<>"New"'], font=Font(name=FONT, bold=True, color=PRE_DARK), fill=fill(PRE_LIGHT)))
lg.conditional_formatting.add(f"M{LOG_FIRST}:M{LOG_LAST}", FormulaRule(
    formula=[f'$M{LOG_FIRST}<>""'], font=Font(name=FONT, bold=True, color="7F6000"), fill=fill("FFF2CC")))
lg.page_setup.orientation = "landscape"
lg.page_setup.fitToWidth = 1
lg.page_setup.fitToHeight = 0
lg.sheet_properties.pageSetUpPr.fitToPage = True

# =============================================================================
# MONTHLY SUMMARY
# =============================================================================
sm = wb.create_sheet("Monthly Summary")
sm.sheet_properties.tabColor = ACCENT
sm.sheet_view.showGridLines = False
banner(sm, "A1:O1", "A2:O2", "MONTHLY SUMMARY",
       "Students, equivalent sales, royalty, marketing, total dues, retention and drops per month")
heads = [("Month", NAVY), ("Toddlers\nStudents", TOD_DARK), ("Preschool\nStudents", PRE_DARK),
         ("One-on-One\nStudents", TUT_DARK), ("Total\nStudents", NAVY), ("Toddlers\nPer Month", TOD_DARK),
         ("Preschool\nPer Month", PRE_DARK), ("One-on-One\nPer Month", TUT_DARK), ("Equivalent\nSales", NAVY),
         ("Royalty", NAVY), ("Marketing\nContribution", NAVY), ("Total Dues", NAVY), ("Retained", NAVY),
         ("Retention\nRate", NAVY), ("Dropped", NAVY)]
for j, (h, col) in enumerate(heads):
    sm.column_dimensions[L(j + 1)].width = 13 if j else 14
    header_cell(sm.cell(4, j + 1), h, color=col)
sm.row_dimensions[4].height = 34
for i, m in enumerate(MONTHS):
    r = 5 + i
    vals = [m, f"=COUNTA({m}!C{FIRST_ROW}:C{LAST_ROW})", f"=COUNTA({m}!K{FIRST_ROW}:K{LAST_ROW})",
            f"=COUNTA({m}!S{FIRST_ROW}:S{LAST_ROW})", f"={m}!B5", f"={m}!G8", f"={m}!O8", f"={m}!V8",
            f"={m}!D5", f"={m}!F5", f"={m}!I5", f"={m}!L5", f"={m}!Z4", f"={m}!O5", f"={m}!T5"]
    for j, v in enumerate(vals):
        c = sm.cell(r, j + 1, v)
        c.border = BORDER
        c.fill = fill(TUT_LIGHT) if i % 2 else fill(WHITE)
        c.font = font(10, j in (0, 4, 8, 11), "000000")
        if j in (5, 6, 7, 8, 9, 10, 11):
            c.number_format, c.alignment = PESO, RIGHT
        elif j == 13:
            c.number_format, c.alignment = PCT, MID
        elif j:
            c.alignment = MID
r = 17
sm.cell(r, 1, "TOTAL / AVG")
for j in range(2, 16):
    col = L(j)
    if j == 14:
        f = f'=IFERROR(AVERAGE({col}5:{col}16),"N/A")'
    else:
        f = f"=SUM({col}5:{col}16)"
    sm.cell(r, j, f)
for j in range(1, 16):
    c = sm.cell(r, j)
    c.font, c.fill, c.border = font(10, True, ACCENT), fill(NAVY), BORDER
    c.number_format = PESO if j in (6, 7, 8, 9, 10, 11, 12) else (PCT if j == 14 else "0")
    c.alignment = RIGHT if j in (6, 7, 8, 9, 10, 11, 12) else MID
sm["A17"].alignment = LEFT
sm["A19"] = ("Equivalent Sales = Σ (Price ÷ package months) of every student on that month's sheet.  "
             "Royalty and Marketing use the rates on the Settings sheet; Total Dues = Royalty + Marketing.  Retention Rate = students who were also "
             "enrolled last month ÷ total students last month × 100 (January has no previous month).  "
             "Students TOTAL counts student-months; Retention 'TOTAL / AVG' is the average of the monthly rates.")
sm["A19"].font = font(9, False, GREY_TXT, True)
sm["A19"].alignment = Alignment(wrap_text=True, vertical="top")
sm.merge_cells("A19:O20")
sm.row_dimensions[19].height = 24
sm["A22"] = "Enrollment records"
sm["A22"].font = font(10, True, NAVY)
for k, (lbl, f) in enumerate([
        ("Active", f"=COUNTIF('Enrollment History'!$R${LOG_FIRST}:$R${LOG_LAST},\"Active\")"),
        ("Dropped", f"=COUNTIF('Enrollment History'!$R${LOG_FIRST}:$R${LOG_LAST},\"Dropped\")"),
        ("Cancelled", f"=COUNTIF('Enrollment History'!$R${LOG_FIRST}:$R${LOG_LAST},\"Cancelled\")"),
        ("Re-enrollments", f"=COUNTIF('Enrollment History'!$D${LOG_FIRST}:$D${LOG_LAST},\"Re-enrollment\")"),
        ("Carry-overs (from last year)", f"=COUNTIF('Enrollment History'!$D${LOG_FIRST}:$D${LOG_LAST},\"Carry-over\")"),
        ("Total enrollments", f"=COUNTA('Enrollment History'!$A${LOG_FIRST}:$A${LOG_LAST})"),
        ("Previous-year records", "=COUNTA('Previous Years'!$A$5:$A$3004)")]):
    sm.cell(23 + k, 1, lbl).font = font(10)
    c = sm.cell(23 + k, 2, f)
    c.font, c.alignment = font(10, True), MID

ch = BarChart()
ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
ch.title = "Equivalent Sales per Month (₱)"
ch.height, ch.width = 8.5, 24
for col_idx in (6, 7, 8):
    ch.add_data(Reference(sm, min_col=col_idx, min_row=4, max_row=16), titles_from_data=True)
ch.set_categories(Reference(sm, min_col=1, min_row=5, max_row=16))
for s, colr in zip(ch.series, [TOD_DARK, PRE_DARK, TUT_DARK]):
    s.graphicalProperties.solidFill = colr
    s.graphicalProperties.line.solidFill = colr
ch.roundedCorners = False
ch.legend.position = "b"
ch.legend.overlay = False
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.y_axis.numFmt = '"₱"#,##0'
sm.add_chart(ch, "A29")
sm.freeze_panes = "B5"

# =============================================================================
# STUDENT PROFILE / SEARCH (written by VBA)
# =============================================================================
ss = wb.create_sheet("Student Search")
ss.sheet_properties.tabColor = PRE_DARK
ss.sheet_view.showGridLines = False
banner(ss, "A1:M1", "A2:M2", "STUDENT PROFILE", "Every enrollment of the searched student(s) — this year and previous years")
ss.row_dimensions[3].height = 26
ss["A3"] = "Search:"
ss["A3"].font = font(10, True, NAVY)
ss["B3"].font = font(11, True, INPUT_TXT)
ss["D3"].font = font(10, True, GREY_TXT)
SH = [("Year", 8), ("Student ID", 14), ("Full Name", 26), ("Enrollment ID", 15), ("Type", 13), ("Program", 22),
      ("Package", 11), ("Start Month", 11), ("Months Covered", 40), ("Net Price", 12), ("Per Month", 12),
      ("Status", 10), ("Last Month", 11)]
for j, (h, w) in enumerate(SH):
    ss.column_dimensions[L(j + 1)].width = w
    header_cell(ss.cell(5, j + 1), h, color=PRE_DARK)
ss.row_dimensions[5].height = 22
for r in range(6, 506):
    for j in range(1, 14):
        c = ss.cell(r, j)
        c.font = font(9)
        if j in (10, 11):
            c.number_format = PESO
ss.conditional_formatting.add("A6:M505", FormulaRule(formula=['$L6="Dropped"'], font=Font(name=FONT, color="C00000")))
ss.conditional_formatting.add("A6:M505", FormulaRule(formula=['$L6="Cancelled"'], font=Font(name=FONT, color="A6A6A6", strike=True)))
ss.conditional_formatting.add("A6:M505", FormulaRule(formula=['AND($A6<>"",MOD(ROW(),2)=1)'], fill=fill(PRE_LIGHT)))
ss.freeze_panes = "A6"

# =============================================================================
# YEAR TRANSFER: END OF YEAR (out) / BEGINNING OF YEAR (in) / PREVIOUS YEARS (archive)
# Same columns as Enrollment History + Record Year + Carry to Year.
# =============================================================================
REC_COLS = HIST + [("Record Year", 9), ("Carry to Year", 9)]
CARRY_HDR, CARRY_FIRST, CARRY_LAST = 8, 9, 3008
ARCH_FIRST, ARCH_LAST = 5, 3004


def style_record_rows(ws, first, last, ncol, banded_rows=300):
    for r in range(first, last + 1):
        for j in range(1, ncol + 1):
            c = ws.cell(r, j)
            if j in (14, 15, 16, 17):
                c.number_format = PESO
            elif j in (3, 21):
                c.number_format = "yyyy-mm-dd hh:mm"
            elif r - first < banded_rows:
                c.font = font(9)


def build_carry(ws, title, sub, note, with_result):
    ws.sheet_view.showGridLines = False
    heads = REC_COLS + ([("Import Result", 40)] if with_result else [])
    ncol = len(heads)
    banner(ws, f"A1:{L(ncol)}1", f"A2:{L(ncol)}2", title, sub)
    ws.merge_cells("A3:N3")
    ws["A3"] = note
    ws["A3"].font = font(9, True, "7F6000")
    ws["A3"].fill = fill("FFF2CC")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    ws.row_dimensions[3].height = 54
    rng = lambda col: f"${col}${CARRY_FIRST}:${col}${CARRY_LAST}"
    ws["A4"] = "Remaining months →"
    ws["A5"] = "Students carried"
    ws["A6"] = "Equivalent sales"
    for r in (4, 5, 6):
        c = ws.cell(r, 1)
        c.font, c.fill, c.border, c.alignment = font(9, True, WHITE), fill(NAVY), BORDER, LEFT
    for i, m in enumerate(MONTHS):
        col = L(2 + i)
        h = ws.cell(4, 2 + i, m)
        h.font, h.fill, h.border, h.alignment = font(9, True, ACCENT), fill(NAVY), BORDER, MID
        c = ws.cell(5, 2 + i, f'=COUNTIFS({rng("M")},"*"&{col}$4&"*",{rng("X")},">0")')
        c.font, c.border, c.alignment = font(10, True, NAVY), BORDER, MID
        c = ws.cell(6, 2 + i, f'=SUMIFS({rng("Q")},{rng("M")},"*"&{col}$4&"*",{rng("X")},">0")')
        c.font, c.border, c.alignment, c.number_format = font(9), BORDER, RIGHT, PESO
    for j, (lbl, f) in enumerate([("Records", f"=COUNTA({rng('A')})"),
                                  ("Remaining packages", f'=COUNTIF({rng("X")},">0")')]):
        c = ws.cell(4, 14 + j, lbl)
        c.font, c.fill, c.border, c.alignment = font(9, True, ACCENT), fill(NAVY), BORDER, CENTER
        c = ws.cell(5, 14 + j, f)
        c.font, c.border, c.alignment = font(12, True, NAVY), BORDER, MID
    for j, (h, w) in enumerate(heads):
        ws.column_dimensions[L(j + 1)].width = max(w, 10)
        header_cell(ws.cell(CARRY_HDR, j + 1), h, color=TOD_DARK if with_result else PRE_DARK)
    ws.column_dimensions["A"].width = 18
    ws.row_dimensions[CARRY_HDR].height = 30
    style_record_rows(ws, CARRY_FIRST, CARRY_LAST, ncol)
    ws.conditional_formatting.add(f"A{CARRY_FIRST}:{L(ncol)}{CARRY_LAST}", FormulaRule(
        formula=[f'$X{CARRY_FIRST}>0'], fill=fill("FFF2CC"), font=Font(name=FONT, bold=True, color="7F6000")))
    ws.freeze_panes = f"B{CARRY_FIRST}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


eoy = wb.create_sheet("End of Year")
eoy.sheet_properties.tabColor = "7F6000"
build_carry(eoy, "END OF YEAR — FULL RECORD FOR NEXT YEAR",
            "Every enrollment and student profile of this file (filled automatically — do not type here)",
            "At year end: click COPY FOR NEXT YEAR → open NEXT year's file → sheet 'Beginning of Year' → click A9 → "
            "press Ctrl+V. Next year's file then takes over everything automatically: all student profiles and "
            "enrollment history, and the remaining months of unfinished packages (highlighted rows, e.g. an "
            "October–March package continues in January–March).", False)
boy = wb.create_sheet("Beginning of Year")
boy.sheet_properties.tabColor = "7F6000"
build_carry(boy, "BEGINNING OF YEAR — CONTINUATION FROM LAST YEAR",
            "Paste last year's End of Year record at A9 — the transfer starts automatically",
            "Click cell A9 and press Ctrl+V (paste what you copied from LAST year's End of Year). The transfer runs "
            "automatically: profiles and history go to 'Previous Years', and unfinished packages are added to this "
            "year's monthly sheets as Carry-over. Pasting the same record again is safe — nothing is added twice. "
            "(If macros were off when you pasted, click IMPORT / TRANSFER.)", True)

arch = wb.create_sheet("Previous Years")
arch.sheet_properties.tabColor = "808080"
arch.sheet_view.showGridLines = False
banner(arch, f"A1:{L(len(REC_COLS))}1", f"A2:{L(len(REC_COLS))}2", "PREVIOUS YEARS",
       "Student profiles and enrollments transferred from earlier years (read-only record)")
for j, (h, w) in enumerate(REC_COLS):
    arch.column_dimensions[L(j + 1)].width = w
    header_cell(arch.cell(4, j + 1), h, color="595959")
arch.row_dimensions[4].height = 30
style_record_rows(arch, ARCH_FIRST, ARCH_LAST, len(REC_COLS))
arch.freeze_panes = "C5"
arch.auto_filter.ref = f"A4:{L(len(REC_COLS))}{ARCH_LAST}"
name("ArchSID", f"'Previous Years'!$B${ARCH_FIRST}:$B${ARCH_LAST}")
name("ArchName", f"'Previous Years'!$G${ARCH_FIRST}:$G${ARCH_LAST}")

order = ["Dashboard"] + MONTHS + ["Monthly Summary", "Enrollment History", "Student Search",
                                  "End of Year", "Beginning of Year", "Previous Years", "Settings"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = ws.title == "Dashboard"

# VBA code names (ThisWorkbook / Sheet1..SheetN) - used by build_xlsm.py
wb.code_name = "ThisWorkbook"
for i, ws in enumerate(wb.worksheets, start=1):
    ws.sheet_properties.codeName = f"Sheet{i}"

wb.save(OUT)
print("saved", OUT)
