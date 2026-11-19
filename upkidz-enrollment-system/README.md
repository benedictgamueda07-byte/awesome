# UpKidz Philippines — Enrollment Management System (Excel + VBA) · v3

| File | What it is |
|---|---|
| **`UpKidz_Enrollment_System.xlsm`** | **Use this one (School Year 2026).** A ready-to-use workbook with the macros and buttons built in. It already contains your record UPK-0001 |
| `UpKidz_Enrollment_System_2027.xlsm` | A blank file for **next year** (School Year 2027). Paste this year's End of Year into it to continue |
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

## 2. Features
| Request | How it works now |
|---|---|
| Package = duration | Duration was removed. The **Package** decides how many months are covered (Settings → *Months*). |
| Automatic Program Fee | Comes from the **Settings price list** for the chosen Program + Package. Only the Discount is typed. |
| Schedule removed | Gone from the form, the monthly sheets, and the history. |
| Re-enrollment | Type the name in **Student** → **RE-ENROLL**. The form fills in the name and the next start month; you choose the new program/package → **SUBMIT**. It is saved as a *Re-enrollment* under the same **Student ID**. Any submit with a name already in the history is automatically treated as a re-enrollment. |
| Student profile / history | **PROFILE / SEARCH** lists every enrollment of the student (programs, months, status, drops) on the *Student Search* sheet. |
| Dropped students | **LOAD FOR EDIT** → set **Enrollment Status = Dropped** and the **Last Month Attended** → **UPDATE ENROLLMENT**. Example: a 6-month Toddler package from January, last month March → the student stays in January–March and is removed from April, May and June. The history keeps Status *Dropped* plus the last month. |
| Computation part on every month | A KPI band at the top of each month shows **Total Students · Equivalent Sales · Royalty (5%) · Marketing Contribution (2%) · TOTAL DUES (Royalty + Marketing) · Retention Rate · Dropped**. |
| Year-end continuation | Packages that run past December (e.g. October → March) are **not** put in this year's January–March. At year end, copy **End of Year**, then paste it into next year's **Beginning of Year**. All profiles, history and the remaining package months transfer automatically (see section 3). |
| Year-stamped IDs | New IDs look like **UPK-2026-0001** (enrollment) and **STU-2026-0001** (student), so they never repeat across yearly workbooks. Your existing UPK-0001 / STU-0001 stays as it is. |
| Snapshot and summary | The Dashboard **Monthly Snapshot** and the **Monthly Summary** sheet show all of these per month. |
| Colors | #375623 / #E2EFDA Toddlers, #1F4E78 / #DDEBF7 Preschool, #203764 / #D9E1F2 One-on-One and headers, #FFC107 accents. |
| CLEAR FORM fixed | It now works even if a form cell has been removed. |
| Sorted lists | Every program table on every month is sorted **longest package first** (10 → 6 → 3 → 1 months; VIP / 2x / 4x count by their months), then **names A–Z**. It re-sorts automatically after every submit, update, drop, cancel and year transfer (macro: `SortAllMonths`). |
| Capitalised names | Names are saved as "Benedict Gamueda" however they are typed. This covers the form, carried-over records, and names typed straight into a monthly sheet's Full Name column. Hyphens and apostrophes are handled ("Maria O'Neil-Santos"). |
| Per-Month Equivalent | Removed from the form. It is still shown per student (**Per Month** column) on the monthly sheets. |

## 3. Year-end: continuing into next year's file
Next year's file continues where this year's file stopped: every student profile and enrollment history
comes over, and every unfinished package continues in the new year's monthly sheets.

1. **At year end, in this year's file:** go to **End of Year** and click **COPY FOR NEXT YEAR**.
   The sheet always holds the full record (refreshed automatically): every enrollment of this year,
   plus everything that came from earlier years. Highlighted rows are unfinished packages; the top shows
   how many students and how much equivalent sales continue in each month.
2. **In next year's file** (`UpKidz_Enrollment_System_2027.xlsm`, or any copy with Settings → School Year
   set to the new year): open **Beginning of Year**, click **A9**, and press **Ctrl+V**.
3. **The transfer runs by itself** as soon as you paste:
   - **Student profiles and history** → saved on the **Previous Years** sheet. PROFILE / SEARCH shows
     them (with the year), the form recognises returning students ("Returning student · STU-2026-0001"),
     and RE-ENROLL keeps the same Student ID.
   - **Unfinished packages** (e.g. October–March) → added to the new year's monthly sheets for the
     remaining months (January–March) and to Enrollment History as **Carry-over**, with their original
     Enrollment ID, Student ID and price. Drops recorded last year are respected.
   - The **Import Result** column shows what happened to each row. Pasting again is safe, because
     nothing is added twice.
4. It keeps going year after year: next year's End of Year includes the Previous Years records, so 2028
   gets 2026 and 2027, and only the newest unfinished packages continue.

> If macros were disabled when you pasted, enable them and click **IMPORT / TRANSFER** on the Beginning
> of Year sheet; it does the same thing. If you changed prices or programs, copy the **Settings** sheet's
> lists into the new file as well: a carried program must exist there.

## 4. Formulas
- **Per Month** (per student) = Net Price ÷ package months. Example: ₱18,000 for 6 Months = ₱3,000.
- **Equivalent Sales** (per month) = sum of *Per Month* for every student on that month's sheet.
- **Royalty** = Equivalent Sales × 5%. **Marketing Contribution** = Equivalent Sales × 2%.
  Both rates can be changed on the Settings sheet (B3, B4).
- **Retention Rate** = students who were also enrolled last month ÷ total students last month × 100.
  Students are matched by name across all programs, so moving from Toddler to Preschool still counts
  as retained. January shows N/A because it has no previous month in the workbook.
- **Total Dues** = Royalty + Marketing Contribution.
- **Dropped** = enrollments whose Last Month Attended is that month (in that school year).

## 5. Settings sheet
- **Price list (A7:D40):** Program, Package, Months, Program Fee. Keep each program's rows together.
- **Programs (G22:H31):** each program and the monthly-sheet section it goes to (Toddlers Group,
  Preschool Group, or One-on-One Tutoring). Toddler PlayTrack VIP → Toddlers Group, and
  Tutoring 4x / 2x → One-on-One. On the monthly sheets the variant is shown in the Package column,
  e.g. *VIP 6 Months* or *2x 10 Months*.
- **Royalty and marketing rates (B3, B4)** and the **School Year (B2)**.

## 6. Monthly sheet columns
Toddlers / Preschool: **No. · Full Name · Package · Discount · Price · Per Month**.
One-on-One: **No. · Full Name · Package · Price · Per Month**. *Price* is the net package price
(fee − discount). Each section shows its package-value total and per-month total in the top-right
corner. A hidden column next to *Per Month* holds the Enrollment ID, so don't delete it.

## 7. Testing guide (fresh workbook)
| # | Do this | Expected |
|---|---|---|
| 1 | Toddler PlayTrack · Benedict Gamueda · 6 Months · January · discount 2000 | UPK-2026-0001 / STU-2026-0001 in January–June. Price ₱16,000, Per Month ₱2,666.67 |
| 2 | Load UPK-2026-0001 → Dropped, Last Month **June** → UPDATE | Blocked: June is the package's final month, so no drop is needed |
| 3 | Same, Last Month **March** → UPDATE | Stays in January–March, removed from April–June. March shows Dropped = 1 |
| 4 | Student `Benedict` → RE-ENROLL → Preschool PlayTrack · 3 Months → SUBMIT | Start Month pre-filled as April. Saved as a Re-enrollment under STU-2026-0001 |
| 5 | Preschool PlayTrack · Ana Cruz · 6 Months · **October** | Appears October–December only. **January–March are NOT filled**; she is on End of Year with "January, February, March" |
| 6 | Check a month's KPIs | January (Benedict only): Equivalent ₱2,666.67, Royalty ₱133.33, Marketing ₱53.33, **Total Dues ₱186.67** |
| 7 | End of Year → COPY FOR NEXT YEAR → 2027 file → Beginning of Year → A9 → Ctrl+V | Runs automatically: 6 profiles saved to Previous Years, Ana / Miguel / Liza continue in 2027. Pasting again adds nothing |
| 7b | In 2027: Student `Benedict` → PROFILE, then type his name in the form | 2026 enrollments are listed. The form shows "Returning student · STU-2026-0001" |
| 8 | In 2027: load Ana's ID → Dropped, Last Month February → UPDATE | Removed from March 2027. Choosing November is blocked, because that month belongs to last year |

## 8. Maintenance
Edit `UpKidzEnrollment.bas` and/or `build_workbook.py`, then run `python3 build_workbook.py`,
then `python3 build_xlsm.py` (standard library plus `openpyxl`). This produces an **empty**
workbook, so copy existing data across or ask for a migration. Keep the `.bas` ASCII-only.
