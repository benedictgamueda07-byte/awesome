# UpKidz Philippines — Enrollment Management System (Excel + VBA) · v2

| File | What it is |
|---|---|
| **`UpKidz_Enrollment_System.xlsm`** | **Use this one.** A ready-to-use workbook with the macros and buttons built in. It already contains your record UPK-0001, carried over from v1 |
| `UpKidz_Enrollment_System.xlsx` | The same workbook without macros (fallback) |
| `UpKidzEnrollment.bas` | The VBA source (the same code that is inside the `.xlsm`) |
| `build_workbook.py` / `build_xlsm.py` | Rebuild the workbook and the `.xlsm` (maintenance only; a rebuild starts empty) |

Requires **Microsoft Excel Desktop for Windows** (Excel 2010 or later).

## 1. Getting started
1. Right-click `UpKidz_Enrollment_System.xlsm` → **Properties** → tick **Unblock** → **OK**.
2. Open it and click **Enable Content**.
3. Use the buttons on the Dashboard. If a button ever goes missing, press **Alt + F8** and run
   **InstallUpKidzSystem**.

> **Don't delete rows or columns on the Dashboard or the monthly sheets.** The macros find cells by
> name and position, and deleting them causes `#REF!` errors (this is what broke CLEAR FORM in v1).
> To change fields or columns, ask for a layout change, and the macro will be updated to match.

## 2. What's new in v2
| Request | How it works now |
|---|---|
| Package = duration | Duration was removed. The **Package** decides how many months are covered (Settings → *Months*). |
| Automatic Program Fee | Comes from the **Settings price list** for the chosen Program + Package. Only the Discount is typed. |
| Schedule removed | Gone from the form, the monthly sheets, and the history. |
| Re-enrollment | Type the name in **Student** → **RE-ENROLL**. The form fills in the name and the next start month; you choose the new program/package → **SUBMIT**. It is saved as a *Re-enrollment* under the same **Student ID**. Any submit with a name already in the history is automatically treated as a re-enrollment. |
| Student profile / history | **PROFILE / SEARCH** lists every enrollment of the student (programs, months, status, drops) on the *Student Search* sheet. |
| Dropped students | **LOAD FOR EDIT** → set **Enrollment Status = Dropped** and the **Drop Month** → **UPDATE ENROLLMENT**. The student is removed from the Drop Month onward, and the history keeps Status *Dropped* plus the Drop Month. |
| Computation part on every month | A KPI band at the top of each month shows **Total Students · Equivalent Sales · Royalty (5%) · Marketing Contribution (2%) · Retention Rate · Dropped**. |
| Snapshot and summary | The Dashboard **Monthly Snapshot** and the **Monthly Summary** sheet show all of these per month. |
| Colors | #375623 / #E2EFDA Toddlers, #1F4E78 / #DDEBF7 Preschool, #203764 / #D9E1F2 One-on-One and headers, #FFC107 accents. |
| CLEAR FORM fixed | It now works even if a form cell has been removed. |

## 3. Formulas
- **Per Month** (per student) = Net Price ÷ package months. Example: ₱18,000 for 6 Months = ₱3,000.
- **Equivalent Sales** (per month) = sum of *Per Month* for every student on that month's sheet.
- **Royalty** = Equivalent Sales × 5%. **Marketing Contribution** = Equivalent Sales × 2%.
  Both rates can be changed on the Settings sheet (B3, B4).
- **Retention Rate** = students who were also enrolled last month ÷ total students last month × 100.
  Students are matched by name across all programs, so moving from Toddler to Preschool still counts
  as retained. January shows N/A because it has no previous month in the workbook.
- **Dropped** = enrollments whose Drop Month is that month.

## 4. Settings sheet
- **Price list (A7:D40):** Program, Package, Months, Program Fee. Keep each program's rows together.
- **Programs (G22:H31):** each program and the monthly-sheet section it goes to (Toddlers Group,
  Preschool Group, or One-on-One Tutoring). Toddler PlayTrack VIP → Toddlers Group, and
  Tutoring 4x / 2x → One-on-One. On the monthly sheets the variant is shown in the Package column,
  e.g. *VIP 6 Months* or *2x 10 Months*.
- **Royalty and marketing rates (B3, B4)** and the **School Year (B2)**.

## 5. Monthly sheet columns
Toddlers / Preschool: **No. · Full Name · Package · Discount · Price · Per Month**.
One-on-One: **No. · Full Name · Package · Price · Per Month**. *Price* is the net package price
(fee − discount). Each section shows its package-value total and per-month total in the top-right
corner. A hidden column next to *Per Month* holds the Enrollment ID, so don't delete it.

## 6. Testing guide (fresh workbook)
| # | Do this | Expected |
|---|---|---|
| 1 | Toddler PlayTrack · Benedict Gamueda · 6 Months · January · discount 2000 | UPK-0001 / STU-0001, Jan–Jun Toddlers, Price ₱16,000, Per Month ₱2,666.67 |
| 2 | Student = `bened` → RE-ENROLL → Preschool PlayTrack · 3 Months → SUBMIT | Form pre-filled with the name and Start Month July. Saved as UPK-0002 **Re-enrollment** under STU-0001, Jul–Sep Preschool |
| 3 | Toddler PlayTrack VIP · Ana Cruz · 3 Months · February | Toddlers Group, Package shown as "VIP 3 Months", ₱15,000 |
| 4 | One-on-One Tutoring 2x · Miguel Santos · 10 Months · May | May → February (wraps the year), ₱1,570 per month |
| 5 | Ana's ID → LOAD FOR EDIT → Status Dropped, Drop Month March → UPDATE | Ana stays only in February. March shows Dropped = 1. History Status = Dropped |
| 6 | PROFILE / SEARCH `STU-0001` | 2 enrollments: Toddler (New) and Preschool (Re-enrollment) |
| 7 | Check the KPIs | January Equivalent Sales ₱4,236.67, Royalty ₱211.83, Marketing ₱84.73. February retention 100% (2 of 2). March 33.3% (1 of 3) |
| 8 | Status Dropped on a *new* enrollment, Drop Month outside the package, discount > fee, missing name, package not listed for the program | Each one is blocked with a message |

## 7. Maintenance
Edit `UpKidzEnrollment.bas` and/or `build_workbook.py`, then run `python3 build_workbook.py`,
then `python3 build_xlsm.py` (standard library plus `openpyxl`). This produces an **empty**
workbook, so copy existing data across or ask for a migration. Keep the `.bas` ASCII-only.
