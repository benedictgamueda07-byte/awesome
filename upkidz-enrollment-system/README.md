# UpKidz Philippines — Enrollment Management System (Excel + VBA)

| File | What it is |
|---|---|
| `UpKidz_Enrollment_System.xlsx` | The workbook: Dashboard, 12 monthly sheets, Monthly Summary, Enrollment History, Student Search, Settings |
| `UpKidzEnrollment.bas` | The VBA module: Submit, Edit, Update, Cancel, Search, and button installer |
| `build_workbook.py` | Rebuilds the workbook from scratch (optional, for maintenance) |

Requires **Microsoft Excel Desktop for Windows** (Excel 2010 or later; Microsoft 365 recommended).

---

## 1. Set up (one time, about 2 minutes)

1. Open `UpKidz_Enrollment_System.xlsx` in Excel. Click **Enable Editing** if Excel asks.
2. Press **Alt + F11** to open the VBA editor.
3. Go to **File → Import File…**, choose `UpKidzEnrollment.bas`, and click **Open**.
   A module named `UpKidzEnrollment` appears under *Modules*.
4. Close the VBA editor and go back to Excel.
5. Press **Alt + F8**, select **InstallUpKidzSystem**, and click **Run**.
   This creates all the buttons on the Dashboard and the Student Search sheet.
6. Go to **File → Save As**, choose **Excel Macro-Enabled Workbook (\*.xlsm)**, and save as
   `UpKidz_Enrollment_System.xlsm`. From now on, use the `.xlsm` file.
7. When you open the `.xlsm` file later, click **Enable Content** on the yellow security bar.
   If Excel blocks macros entirely, right-click the file → **Properties** → tick **Unblock** → OK.

### Creating or reassigning the buttons manually (only if you need to)

`InstallUpKidzSystem` does this for you. To do it yourself:
**Insert → Shapes → Rounded Rectangle**, draw it, type the label, then right-click it →
**Assign Macro…** → choose the macro:

| Button label | Macro | Suggested place |
|---|---|---|
| SUBMIT ENROLLMENT | `SubmitEnrollment` | Dashboard B18:C18 |
| UPDATE ENROLLMENT | `UpdateEnrollment` | Dashboard D18 |
| CLEAR FORM | `ClearForm` | Dashboard E18 |
| LOAD FOR EDIT | `LoadEnrollmentForEdit` | Dashboard G8:H8 |
| CANCEL ENROLLMENT | `CancelEnrollment` | Dashboard I8:J8 |
| SEARCH STUDENT | `SearchStudent` | Dashboard G12:J12 |
| BACK TO DASHBOARD | `GoToDashboard` | Student Search H3:J3 |

---

## 2. How the system works

### Dashboard → one click → every covered month
1. Fill in the **yellow cells** (blue text marks your inputs): Program, Full Name, Schedule,
   Package, Start Month, Duration, Program Fee, and Discount.
2. **Duration Display** (e.g. `January–June (6 months)`), **Net Price**, and **Status Check**
   calculate automatically. Hints in column E show the package length and the standard fee
   from Settings.
3. Click **SUBMIT ENROLLMENT**. The macro:
   - validates everything (see the list below) and shows a confirmation summary,
   - works out the covered months, wrapping past December (November + 3 → Nov, Dec, Jan),
   - checks **every** target month for duplicates and free rows **before writing anything**, so
     an enrollment is never half-saved,
   - writes the student to the next empty row of the correct program section only, in each month,
   - gives the enrollment an ID (`UPK-0001`, `UPK-0002`, …) and records it in **Enrollment History**,
   - clears the form, which also stops the same enrollment being submitted twice, and shows a
     success message.

### Monthly sheets
Each month shows **Toddlers Group** (green), **Preschool Group** (blue), and **One-on-One Tutoring**
(dark blue) side by side, with 100 rows each. Every section has its own:
- automatic numbering (`No.` column, formula-driven),
- **TOTAL** in its top-right corner and a **STUDENTS** count just below it,
- alternating row colours, a print setup (A4 landscape, fit to width, header rows repeat), and a
  frozen header.

Toddlers and Preschool record the **Discount** and the **net price** in *Price*. One-on-One has
no Discount column, so its *Price* holds the net price.

Each row also stores its **Enrollment ID** in a hidden column next to *Price*. Edit and Cancel use
this ID to find exactly the right rows in every month, so they never leave stale or duplicate
copies behind. Don't delete those hidden columns.

### Price recorded per month (Settings!B3)
- **Full package price** (default, as in your example): ₱16,700 appears in each of the six months.
  This means monthly totals add up the full price of every student active that month.
- **Split per month**: the net price is divided by the duration (₱16,700 ÷ 6 = ₱2,783.33 per
  month), so the 12 monthly totals add up to the real revenue.

### Edit an enrollment
Type the ID in **Enrollment ID** → **LOAD FOR EDIT** → change any field (including program, start
month, or duration) → **UPDATE ENROLLMENT**. The old rows are removed from every month, the lists
close up, the new rows are written, and the history row is updated (Remarks shows "Edited
yyyy-mm-dd").

### Cancel an enrollment
Type the ID → **CANCEL ENROLLMENT** → confirm. The student is removed from every month. The
history row is kept with Status **Cancelled**, shown greyed out and struck through.

### Search for a student
Type all or part of a name in **Search Student** → **SEARCH STUDENT**. The **Student Search** sheet
lists every month and section where the student appears, with their ID and status.

### Monthly Summary
Shows students and total value per program for each month, grand totals, and a stacked chart. The
Dashboard's **Monthly Snapshot** shows the same figures for whichever month you pick.

### Validation and error prevention
- Program, Full Name, Schedule, Package, Start Month, Duration (whole number 1–12), and Program
  Fee (> 0) are all required.
- A discount larger than the fee is blocked.
- The package must belong to the selected program.
- If the duration doesn't match the package (e.g. a "3 Months" package with 5 months), you are
  asked to confirm. The default answer is No.
- **Duplicate guard:** if the same name (ignoring capitals and extra spaces) is already in that
  section for any target month, you are warned. The default answer is No.
- **Capacity guard:** if any target month's section is full, nothing is saved.
- Rows are only written where *every* cell is empty, so existing records are never overwritten.
- Dropdowns and data validation on the Dashboard catch most mistakes before you even click.

### Settings sheet (easy to maintain)
- **Packages:** program, package name, months, and standard fee. Keep each program's rows together.
  The Package dropdown on the Dashboard filters to the selected program automatically.
  ⚠ Only *Toddler 6 Months = ₱18,700* came from you. **All other fees are sample values; replace
  them with the real UpKidz rates.**
- **Schedules:** dropdown suggestions. You can still type any schedule.
- **School Year:** used in the monthly sheet titles.

> Use the Dashboard to change enrolled students. If you type directly on a monthly sheet, that
> row has no Enrollment ID: Search marks it "Manual entry", and Edit/Cancel can't manage it.

---

## 3. Testing guide (sample enrollments)

Run these in order on a fresh copy. Expected results are what this system produced in testing.

| # | Input | Expected result |
|---|---|---|
| 1 | Toddler PlayTrack · **Benedict Gamueda** · Monday and Wednesday, 9:00–10:00 AM · 6 Months · January · 6 · 18700 · 2000 | ID UPK-0001. In **Toddlers Group** for Jan, Feb, Mar, Apr, May, and Jun only. Discount ₱2,000, Price ₱16,700. Jan Toddlers TOTAL = ₱16,700. Not in Preschool or One-on-One. |
| 2 | Preschool PlayTrack · Ana Cruz · Tue & Thu 9–10 AM · 3 Months · **March** · 3 · 10800 · 500 | UPK-0002. Preschool for **Mar, Apr, May** only. Price ₱10,300. |
| 3 | One-on-One Tutoring · Miguel Santos · Saturday 9–11 AM · 3 Months · **November** · 3 · 12900 · 900 | UPK-0003. One-on-One for **Nov, Dec, Jan** (crosses the year). Price ₱12,000. No discount column. |
| 4 | Toddler PlayTrack · Liza Reyes · 3 Months · February · 3 · 9900 · 0 | UPK-0004. Toddlers for Feb–Apr. In March she is **No. 2**, under Benedict. |
| 5 | Repeat #1 with name `  benedict   gamueda ` starting March | **Possible duplicate** warning listing Mar–Jun (UPK-0001). Answer No → nothing is added. |
| 6 | Preschool · fee 3800 · discount 5000 | Blocked: "discount is greater than the program fee". |
| 7 | Leave Full Name blank | Blocked, with a list of what's missing. |
| 8 | Package 3 Months with Duration 5 | Asked to confirm. Answer No → nothing is added. |
| 9 | ID `UPK-0004` → LOAD FOR EDIT → Start Month April, discount 400 → UPDATE | Liza is removed from Feb and Mar and now appears in **Apr, May, Jun** at ₱9,500. History says "Edited". |
| 10 | ID `UPK-0002` → CANCEL ENROLLMENT | Ana is removed from Mar–May. History Status = Cancelled. Cancelling again says "already cancelled". |
| 11 | Search `gamu` | 6 results (Jan–Jun, Toddlers Group, UPK-0001, Active). |
| 12 | Settings!B3 = Split per month, then Preschool 3 Months · December · 3 · 10800 · 300 | Dec, Jan, and Feb each show ₱3,500.00. |
| 13 | Click SUBMIT while editing | Told to use UPDATE or CLEAR FORM. |

To reset after testing, clear the data cells on the monthly sheets (Full Name through Price, plus
the hidden ID column), clear Enrollment History from row 5 down, and clear Student Search.
Don't clear the **No.** column, which holds formulas.

---

## 4. Maintenance notes
- **More than 100 rows per section:** change `LAST_ROW` at the top of the VBA module and in
  `build_workbook.py`, then rebuild. Alternatively, copy the formatting and the `No.` formula down
  on every month and update `LAST_ROW`.
- **Layout constants** (section start columns, first and last row, sheet names) are at the top of
  the `.bas` module and match `build_workbook.py`.
- **Dashboard cells** are referenced by names (`inProgram`, `inName`, … `inEditID`, `inManageID`,
  `inSearch`), so you can move them without editing code. Use **Formulas → Name Manager**.
- Rebuild a clean workbook with: `python3 build_workbook.py UpKidz_Enrollment_System.xlsx`
  (requires `openpyxl`).
