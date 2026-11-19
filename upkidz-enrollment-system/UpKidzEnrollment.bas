Attribute VB_Name = "UpKidzEnrollment"
'==============================================================================
'  UPKIDZ PHILIPPINES - ENROLLMENT MANAGEMENT SYSTEM  (v2)
'
'  Buttons / public macros:
'    SubmitEnrollment      - add a new enrollment (or re-enrollment) to every
'                            month its package covers
'    UpdateEnrollment      - apply edits, including Status = Dropped + Drop Month
'    ClearForm             - clear the Dashboard form / leave edit mode
'    LoadEnrollmentForEdit - load an Enrollment ID into the form
'    CancelEnrollment      - remove an enrollment entered by mistake
'    SearchStudent         - show a student's profile (all enrollments)
'    ReEnrollStudent       - fill the form for a returning student
'    GoToDashboard         - navigation
'    RefreshEndOfYear      - rebuild the End of Year carry-over list
'    ImportCarryOver       - (next year's workbook) import pasted carry-over rows
'    InstallUpKidzSystem   - redraw all buttons
'    EnsureButtons         - runs on open (re-links / adds buttons)
'
'  Months covered come from the PACKAGE (Settings price list). Each record on
'  a monthly sheet carries its Enrollment ID in a hidden column; every
'  enrollment row in Enrollment History carries a Student ID, so all of a
'  student's enrollments (re-enrollments) stay linked.
'==============================================================================
Option Explicit

'---- layout (must match build_workbook.py) -----------------------------------
Private Const FIRST_ROW As Long = 11
Private Const LAST_ROW As Long = 110
Private Const COL_TODDLER As Long = 2        ' B = "No." of Toddlers Group
Private Const COL_PRESCHOOL As Long = 10     ' J = "No." of Preschool Group
Private Const COL_TUTOR As Long = 18         ' R = "No." of One-on-One Tutoring

Private Const SH_DASH As String = "Dashboard"
Private Const SH_LOG As String = "Enrollment History"
Private Const SH_SEARCH As String = "Student Search"
Private Const LOG_FIRST_ROW As Long = 5
Private Const LOG_LAST_ROW As Long = 1005
Private Const SEARCH_FIRST_ROW As Long = 6
Private Const SEARCH_LAST_ROW As Long = 505

' Enrollment History columns
Private Const LC_EID As Long = 1
Private Const LC_SID As Long = 2
Private Const LC_DATE As Long = 3
Private Const LC_TYPE As Long = 4
Private Const LC_PROGRAM As Long = 5
Private Const LC_SECTION As Long = 6
Private Const LC_NAME As Long = 7
Private Const LC_PACKAGE As Long = 8
Private Const LC_MONTHS As Long = 9
Private Const LC_START As Long = 10
Private Const LC_STARTYEAR As Long = 11
Private Const LC_COVERED As Long = 12       ' months in THIS workbook's year
Private Const LC_NEXTYEAR As Long = 13      ' months that fall in next year
Private Const LC_FEE As Long = 14
Private Const LC_DISCOUNT As Long = 15
Private Const LC_NET As Long = 16
Private Const LC_PERMONTH As Long = 17
Private Const LC_STATUS As Long = 18
Private Const LC_DROP As Long = 19          ' last month attended (dropped)
Private Const LC_DROPYEAR As Long = 20
Private Const LC_UPDATED As Long = 21
Private Const LC_REMARKS As Long = 22

' End of Year / Beginning of Year carry-over sheets
Private Const SH_EOY As String = "End of Year"
Private Const SH_BOY As String = "Beginning of Year"
Private Const CARRY_FIRST As Long = 9
Private Const CARRY_LAST As Long = 208

Private Const APP_TITLE As String = "UpKidz Enrollment System"

'---- current form data (filled by ReadForm) ----------------------------------
Private fProgram As String
Private fSection As Long          ' 1 Toddlers, 2 Preschool, 3 One-on-One
Private fName As String
Private fPackage As String
Private fMonths As Long           ' months the package covers
Private fStartMonth As Long
Private fFee As Double
Private fDiscount As Double
Private fNet As Double
Private fPerMonth As Double
Private fStatus As String         ' "Active" or "Dropped"
Private fDropMonth As Long        ' last month attended (0 = not dropped)
Private fDropPos As Long          ' its position in the package (-1 = none)
Private fDropYear As Long
Private fStartOffset As Long      ' start year - this workbook's school year (0 or -1)

'---- test hook: when True, prompts are auto-answered and messages are stored -
Public gSilent As Boolean
Public gSilentAnswer As VbMsgBoxResult
Public gLastMessage As String

'==============================================================================
'  PUBLIC MACROS
'==============================================================================
Public Sub SubmitEnrollment()
    Dim months As Variant, nextM As Variant, newID As String, sid As String, isReturning As Boolean, msg As String

    If Len(CStr(SafeGet("inEditID"))) > 0 Then
        Notify "You are editing enrollment " & SafeGet("inEditID") & "." & vbCrLf & _
               "Click UPDATE ENROLLMENT to save the changes, or CLEAR FORM to start a new enrollment.", vbExclamation
        Exit Sub
    End If
    If Not ReadForm(False) Then Exit Sub
    months = ThisYearList(fMonths)
    nextM = NextYearList(fMonths)

    If Not CheckDuplicates(months, "") Then Exit Sub
    If Not CheckCapacity(months, "", False) Then Exit Sub

    sid = FindStudentID(fName)
    isReturning = (sid <> "")
    msg = "Please confirm this enrollment:" & vbCrLf & vbCrLf
    If isReturning Then
        msg = msg & "RE-ENROLLMENT of returning student " & sid & vbCrLf & _
              "Previous: " & LastEnrollmentText(sid) & vbCrLf & vbCrLf
    End If
    msg = msg & FormSummary(months, nextM) & vbCrLf & "Submit now?"
    If Ask(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    If Not isReturning Then sid = NextStudentID()
    newID = NextEnrollmentID()
    WriteToMonths months, newID
    LogNew newID, sid, IIf(isReturning, "Re-enrollment", "New"), months, nextM, Now
    RefreshEndOfYear
    Application.ScreenUpdating = True

    ClearInputs
    Notify IIf(isReturning, "Re-enrollment", "Enrollment") & " submitted successfully!" & vbCrLf & vbCrLf & _
           "Enrollment ID: " & newID & vbCrLf & _
           "Student ID: " & sid & vbCrLf & _
           "Student: " & fName & vbCrLf & _
           "Program: " & fProgram & " (" & SectionTitle(fSection) & ")" & vbCrLf & _
           "Added to: " & ListOrNone(months) & _
           IIf(ArrLen(nextM) > 0, vbCrLf & "Next year (End of Year sheet): " & ListOrNone(nextM), ""), vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify "Submission stopped because of an error: " & Err.Description, vbCritical
End Sub

Public Sub ClearForm()
    ClearInputs
    SafeSet "inEditID", ""
End Sub

Public Sub LoadEnrollmentForEdit()
    Dim id As String, r As Long, lg As Worksheet
    id = UCase$(Trim$(CStr(SafeGet("inManageID"))))
    If id = "" Then
        Notify "Type the Enrollment ID (e.g. UPK-2026-0001) in the Enrollment ID box first.", vbExclamation
        Exit Sub
    End If
    r = FindLogRow(id)
    If r = 0 Then
        Notify "Enrollment ID " & id & " was not found in Enrollment History.", vbExclamation
        Exit Sub
    End If
    Set lg = Worksheets(SH_LOG)
    If lg.Cells(r, LC_STATUS).Value = "Cancelled" Then
        Notify "Enrollment " & id & " is cancelled and cannot be edited.", vbExclamation
        Exit Sub
    End If
    ClearInputs
    SafeSet "inProgram", lg.Cells(r, LC_PROGRAM).Value
    SafeSet "inName", lg.Cells(r, LC_NAME).Value
    SafeSet "inPackage", lg.Cells(r, LC_PACKAGE).Value
    SafeSet "inStartMonth", lg.Cells(r, LC_START).Value
    SafeSet "inDiscount", lg.Cells(r, LC_DISCOUNT).Value
    SafeSet "inStatus", IIf(CStr(lg.Cells(r, LC_STATUS).Value) = "Dropped", "Dropped", "Active")
    SafeSet "inDropMonth", lg.Cells(r, LC_DROP).Value
    SafeSet "inEditID", id
    SafeSet "inManageID", id
    Notify "Enrollment " & id & " loaded." & vbCrLf & vbCrLf & _
           "Change any field, then click UPDATE ENROLLMENT." & vbCrLf & _
           "If the student stopped: set Enrollment Status = Dropped and choose the LAST MONTH ATTENDED. " & _
           "The student stays in that month and is removed from the following months.", vbInformation
End Sub

Public Sub UpdateEnrollment()
    Dim id As String, r As Long, months As Variant, nextM As Variant, n As Long, removed As Long, msg As String
    id = CStr(SafeGet("inEditID"))
    If id = "" Then
        Notify "No enrollment is loaded for editing." & vbCrLf & _
               "Type an Enrollment ID and click LOAD FOR EDIT first.", vbExclamation
        Exit Sub
    End If
    r = FindLogRow(id)
    If r = 0 Then
        Notify "Enrollment ID " & id & " no longer exists in Enrollment History.", vbCritical
        Exit Sub
    End If
    If Not ReadForm(True) Then Exit Sub
    n = AttendedCount()
    months = ThisYearList(n)
    nextM = NextYearList(n)
    If ArrLen(months) > 0 Then
        If Not CheckDuplicates(months, id) Then Exit Sub
        If Not CheckCapacity(months, id, False) Then Exit Sub
    End If

    msg = "Update enrollment " & id & "?" & vbCrLf & vbCrLf & FormSummary(months, nextM) & vbCrLf
    If fStatus = "Dropped" Then
        msg = msg & "DROPPED - last month attended: " & MonthSheet(fDropMonth) & " " & fDropYear & "." & vbCrLf & _
              "The student is removed from the months after it." & vbCrLf
    End If
    msg = msg & "The old entries will be replaced in every month."
    If Ask(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    removed = RemoveFromMonths(id)
    If ArrLen(months) > 0 Then WriteToMonths months, id
    LogUpdate r, months, nextM
    RefreshEndOfYear
    Application.ScreenUpdating = True

    ClearInputs
    SafeSet "inEditID", ""
    Notify "Enrollment " & id & " updated" & _
           IIf(fStatus = "Dropped", " (DROPPED - last month attended " & MonthSheet(fDropMonth) & ")", "") & "." & vbCrLf & _
           "Old entries removed: " & removed & vbCrLf & _
           "Now recorded in: " & ListOrNone(months) & _
           IIf(ArrLen(nextM) > 0, vbCrLf & "Next year (End of Year sheet): " & ListOrNone(nextM), ""), vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify "Update stopped because of an error: " & Err.Description, vbCritical
End Sub

Public Sub CancelEnrollment()
    Dim id As String, r As Long, lg As Worksheet, removed As Long
    id = UCase$(Trim$(CStr(SafeGet("inManageID"))))
    If id = "" Then id = CStr(SafeGet("inEditID"))
    If id = "" Then
        Notify "Type the Enrollment ID to cancel in the Enrollment ID box.", vbExclamation
        Exit Sub
    End If
    r = FindLogRow(id)
    If r = 0 Then
        Notify "Enrollment ID " & id & " was not found in Enrollment History.", vbExclamation
        Exit Sub
    End If
    Set lg = Worksheets(SH_LOG)
    If lg.Cells(r, LC_STATUS).Value = "Cancelled" Then
        Notify "Enrollment " & id & " is already cancelled.", vbInformation
        Exit Sub
    End If
    If Ask("Cancel this enrollment (entered by mistake)?" & vbCrLf & vbCrLf & _
           "ID: " & id & vbCrLf & "Student: " & lg.Cells(r, LC_NAME).Value & vbCrLf & _
           "Program: " & lg.Cells(r, LC_PROGRAM).Value & vbCrLf & "Months: " & CoveredText(r) & _
           vbCrLf & vbCrLf & "The student will be removed from every month listed." & vbCrLf & _
           "(If the student stopped attending, use Status = Dropped with UPDATE instead.)", _
           vbYesNo + vbExclamation + vbDefaultButton2) <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    removed = RemoveFromMonths(id)
    lg.Cells(r, LC_STATUS).Value = "Cancelled"
    lg.Cells(r, LC_UPDATED).Value = Now
    lg.Cells(r, LC_REMARKS).Value = Trim$(lg.Cells(r, LC_REMARKS).Value & " Cancelled " & Format$(Now, "yyyy-mm-dd") & ".")
    RefreshEndOfYear
    Application.ScreenUpdating = True

    If CStr(SafeGet("inEditID")) = id Then
        ClearInputs
        SafeSet "inEditID", ""
    End If
    SafeSet "inManageID", ""
    Notify "Enrollment " & id & " cancelled." & vbCrLf & "Entries removed from monthly sheets: " & removed, vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify "Cancellation stopped because of an error: " & Err.Description, vbCritical
End Sub

' Student profile: every enrollment of the matching student(s)
Public Sub SearchStudent()
    Dim term As String, ss As Worksheet, lg As Worksheet, r As Long, outR As Long
    Dim sids As String, nStudents As Long, sid As String

    term = Trim$(CStr(SafeGet("inSearch")))
    If term = "" Then
        Notify "Type a name (or part of it), a Student ID or an Enrollment ID in the Student box.", vbExclamation
        Exit Sub
    End If
    Set ss = Worksheets(SH_SEARCH)
    Set lg = Worksheets(SH_LOG)
    ss.Range(ss.Cells(SEARCH_FIRST_ROW, 1), ss.Cells(SEARCH_LAST_ROW, 12)).ClearContents
    ss.Range("B3").Value = term
    sids = MatchingStudentIDs(term)
    outR = SEARCH_FIRST_ROW
    If sids <> "" Then
        For r = LOG_FIRST_ROW To LOG_LAST_ROW
            If CStr(lg.Cells(r, LC_EID).Value) = "" Then Exit For
            sid = CStr(lg.Cells(r, LC_SID).Value)
            If InStr(1, "|" & sids & "|", "|" & sid & "|", vbTextCompare) > 0 And outR <= SEARCH_LAST_ROW Then
                ss.Cells(outR, 1).Value = sid
                ss.Cells(outR, 2).Value = lg.Cells(r, LC_NAME).Value
                ss.Cells(outR, 3).Value = lg.Cells(r, LC_EID).Value
                ss.Cells(outR, 4).Value = lg.Cells(r, LC_TYPE).Value
                ss.Cells(outR, 5).Value = lg.Cells(r, LC_PROGRAM).Value
                ss.Cells(outR, 6).Value = lg.Cells(r, LC_PACKAGE).Value
                ss.Cells(outR, 7).Value = lg.Cells(r, LC_START).Value
                ss.Cells(outR, 8).Value = CoveredText(r)
                ss.Cells(outR, 9).Value = lg.Cells(r, LC_NET).Value
                ss.Cells(outR, 10).Value = lg.Cells(r, LC_PERMONTH).Value
                ss.Cells(outR, 11).Value = lg.Cells(r, LC_STATUS).Value
                ss.Cells(outR, 12).Value = lg.Cells(r, LC_DROP).Value
                outR = outR + 1
            End If
        Next r
        nStudents = UBound(Split(sids, "|")) + 1
    End If
    If outR = SEARCH_FIRST_ROW Then
        ss.Range("D3").Value = "No student found."
    Else
        ss.Range("D3").Value = nStudents & " student(s), " & (outR - SEARCH_FIRST_ROW) & " enrollment(s)."
    End If
    gLastMessage = CStr(ss.Range("D3").Value)
    If Not gSilent Then
        ss.Activate
        ss.Range("A" & SEARCH_FIRST_ROW).Select
    End If
End Sub

' Fill the form for a returning student so a new program/package can be submitted
Public Sub ReEnrollStudent()
    Dim term As String, sids As String, parts As Variant, sid As String, r As Long
    Dim lg As Worksheet, nm As String, nextM As Long, i As Long, lst As String

    If Len(CStr(SafeGet("inEditID"))) > 0 Then
        Notify "Finish or clear the enrollment you are editing first (UPDATE or CLEAR FORM).", vbExclamation
        Exit Sub
    End If
    term = Trim$(CStr(SafeGet("inSearch")))
    If term = "" Then
        Notify "Type the student's name, Student ID or an Enrollment ID in the Student box, then click RE-ENROLL.", vbExclamation
        Exit Sub
    End If
    sids = MatchingStudentIDs(term)
    If sids = "" Then
        Notify "No student matching '" & term & "' was found in Enrollment History." & vbCrLf & _
               "For a first-time student, fill in the form and click SUBMIT ENROLLMENT.", vbExclamation
        Exit Sub
    End If
    parts = Split(sids, "|")
    If UBound(parts) > 0 Then
        For i = 0 To UBound(parts)
            If i < 10 Then lst = lst & "  " & parts(i) & "  " & StudentName(CStr(parts(i))) & vbCrLf
        Next i
        Notify "More than one student matches '" & term & "':" & vbCrLf & vbCrLf & lst & vbCrLf & _
               "Type the full name or the Student ID, then click RE-ENROLL again.", vbExclamation
        Exit Sub
    End If
    sid = CStr(parts(0))
    r = LastEnrollmentRow(sid)
    Set lg = Worksheets(SH_LOG)
    nm = CStr(lg.Cells(r, LC_NAME).Value)
    nextM = NextMonthAfter(r)

    ClearInputs
    SafeSet "inEditID", ""
    SafeSet "inName", nm
    If nextM > 0 Then SafeSet "inStartMonth", MonthSheet(nextM)
    If Not gSilent Then
        Worksheets(SH_DASH).Activate
        Range("inProgram").Select
    End If
    Notify "Re-enrolling " & nm & " (" & sid & ")." & vbCrLf & vbCrLf & _
           "Previous: " & LastEnrollmentText(sid) & vbCrLf & vbCrLf & _
           "Now choose the new Program and Package" & IIf(nextM > 0, " (Start Month set to " & MonthSheet(nextM) & ")", "") & _
           ", then click SUBMIT ENROLLMENT. It will be saved as a re-enrollment under " & sid & ".", vbInformation
End Sub

Public Sub GoToDashboard()
    Worksheets(SH_DASH).Activate
    Range("inProgram").Select
End Sub

'------------------------------------------------------------------------------
'  BUTTONS
'------------------------------------------------------------------------------
Public Sub InstallUpKidzSystem()
    BuildButtons True
    Worksheets(SH_DASH).Activate
    Notify "UpKidz Enrollment System buttons have been (re)created." & vbCrLf & vbCrLf & _
           "Keep saving the workbook as an Excel Macro-Enabled Workbook (*.xlsm).", vbInformation
End Sub

' Called from ThisWorkbook.Workbook_Open: adds missing buttons and re-links
' every button to its macro. Silent.
Public Sub EnsureButtons()
    Dim wasSaved As Boolean
    On Error Resume Next
    wasSaved = ThisWorkbook.Saved
    BuildButtons False
    ThisWorkbook.Saved = wasSaved
End Sub

Private Sub BuildButtons(recreate As Boolean)
    Dim ds As Worksheet, ss As Worksheet, eo As Worksheet, bo As Worksheet
    Set ds = Worksheets(SH_DASH)
    Set ss = Worksheets(SH_SEARCH)
    Set eo = Worksheets(SH_EOY)
    Set bo = Worksheets(SH_BOY)
    If recreate Then
        DeleteButtons ds
        DeleteButtons ss
        DeleteButtons eo
        DeleteButtons bo
    End If
    If CStr(ds.Range("B19").Value) <> "" Then ds.Range("B19").Value = ""
    PlaceButton ds, ds.Range("B19:C20"), "btnSubmit", "SUBMIT ENROLLMENT", "SubmitEnrollment", RGB(55, 86, 35), 14
    PlaceButton ds, ds.Range("D19:D20"), "btnUpdate", "UPDATE ENROLLMENT", "UpdateEnrollment", RGB(31, 78, 120), 10
    PlaceButton ds, ds.Range("E19:E20"), "btnClear", "CLEAR FORM", "ClearForm", RGB(128, 128, 128), 10
    PlaceButton ds, ds.Range("G8:H8"), "btnLoad", "LOAD FOR EDIT", "LoadEnrollmentForEdit", RGB(31, 78, 120), 10
    PlaceButton ds, ds.Range("I8:J8"), "btnCancel", "CANCEL ENROLLMENT", "CancelEnrollment", RGB(192, 0, 0), 10
    PlaceButton ds, ds.Range("G12:H12"), "btnSearch", "PROFILE / SEARCH", "SearchStudent", RGB(32, 55, 100), 10
    PlaceButton ds, ds.Range("I12:J12"), "btnReenroll", "RE-ENROLL", "ReEnrollStudent", RGB(55, 86, 35), 10
    PlaceButton ss, ss.Range("J3:L3"), "btnBack", "BACK TO DASHBOARD", "GoToDashboard", RGB(32, 55, 100), 9
    PlaceButton eo, eo.Range("O3:Q3"), "btnRefreshEoy", "REFRESH LIST", "RefreshEndOfYear", RGB(31, 78, 120), 10
    PlaceButton bo, bo.Range("O3:R3"), "btnImport", "IMPORT CARRY-OVER", "ImportCarryOver", RGB(55, 86, 35), 11
End Sub

'==============================================================================
'  FORM READING & VALIDATION
'==============================================================================
Private Function ReadForm(isEdit As Boolean) As Boolean
    Dim v As Variant, problems As String, i As Long, r As Long, lg As Worksheet, useHist As Boolean

    ReadForm = False
    fProgram = CleanText(CStr(SafeGet("inProgram")))
    fName = CleanText(CStr(SafeGet("inName")))
    fPackage = CleanText(CStr(SafeGet("inPackage")))
    fStartMonth = MonthIndex(CStr(SafeGet("inStartMonth")))
    fStatus = IIf(LCase$(Trim$(CStr(SafeGet("inStatus")))) = "dropped", "Dropped", "Active")
    fDropMonth = MonthIndex(CStr(SafeGet("inDropMonth")))
    fSection = ProgramSection(fProgram)
    fStartOffset = 0
    fDropPos = -1
    fDropYear = 0
    useHist = False
    If isEdit Then
        r = FindLogRow(CStr(SafeGet("inEditID")))
        If r > 0 Then
            Set lg = Worksheets(SH_LOG)
            If IsNumeric(lg.Cells(r, LC_STARTYEAR).Value) And CStr(lg.Cells(r, LC_STARTYEAR).Value) <> "" Then
                fStartOffset = CLng(lg.Cells(r, LC_STARTYEAR).Value) - SchoolYear()
            End If
            ' same program + package as recorded: keep the original price (price list may have changed)
            If StrComp(CleanText(CStr(lg.Cells(r, LC_PROGRAM).Value)), fProgram, vbTextCompare) = 0 And _
               StrComp(CleanText(CStr(lg.Cells(r, LC_PACKAGE).Value)), fPackage, vbTextCompare) = 0 And _
               IsNumeric(lg.Cells(r, LC_MONTHS).Value) And IsNumeric(lg.Cells(r, LC_FEE).Value) Then
                fMonths = CLng(lg.Cells(r, LC_MONTHS).Value)
                fFee = CDbl(lg.Cells(r, LC_FEE).Value)
                useHist = (fMonths > 0 And fFee > 0)
            End If
        End If
    End If

    If fProgram = "" Then
        problems = problems & "- Select a Program." & vbCrLf
    ElseIf fSection = 0 Then
        problems = problems & "- Program '" & fProgram & "' is not in the Settings Programs table (or has no Section)." & vbCrLf
    End If
    If fName = "" Then problems = problems & "- Enter the student's Full Name." & vbCrLf
    If fPackage = "" Then problems = problems & "- Select a Package." & vbCrLf
    If fStartMonth = 0 Then problems = problems & "- Select a Start Month (January to December)." & vbCrLf

    v = SafeGet("inDiscount")
    If IsEmpty(v) Or CStr(v) = "" Then
        fDiscount = 0
    ElseIf IsNumeric(v) Then
        fDiscount = CDbl(v)
        If fDiscount < 0 Then problems = problems & "- Discount cannot be negative." & vbCrLf
    Else
        problems = problems & "- Discount must be a number (leave blank if none)." & vbCrLf
    End If

    If problems = "" And Not useHist Then
        If Not LookupPackage(fProgram, fPackage, fMonths, fFee) Then
            problems = problems & "- Package '" & fPackage & "' is not listed for " & fProgram & " on the Settings sheet." & vbCrLf
        ElseIf fMonths < 1 Or fMonths > 12 Then
            problems = problems & "- Package '" & fPackage & "' must cover 1 to 12 months (Settings, Months column)." & vbCrLf
        ElseIf fFee <= 0 Then
            problems = problems & "- Package '" & fPackage & "' has no Program Fee on the Settings sheet." & vbCrLf
        End If
    End If

    If fStatus = "Dropped" Then
        If Not isEdit Then
            problems = problems & "- Status 'Dropped' is only used when updating an enrollment. " & _
                       "Set Enrollment Status to Active (or blank) for a new enrollment." & vbCrLf
        ElseIf fDropMonth = 0 Then
            problems = problems & "- Select the Last Month Attended." & vbCrLf
        End If
    Else
        fDropMonth = 0
    End If

    If problems <> "" Then
        Notify "The enrollment cannot be saved yet:" & vbCrLf & vbCrLf & problems, vbExclamation
        Exit Function
    End If
    If fDiscount > fFee Then
        Notify "The discount (" & Peso(fDiscount) & ") is greater than the program fee (" & _
               Peso(fFee) & ")." & vbCrLf & "Please correct the discount.", vbExclamation
        Exit Function
    End If
    fNet = fFee - fDiscount
    fPerMonth = Round(fNet / fMonths, 2)

    If fDropMonth > 0 Then
        For i = 0 To fMonths - 1
            If SeqMonth(i) = fDropMonth Then fDropPos = i
        Next i
        If fDropPos < 0 Then
            Notify "Last Month Attended (" & MonthSheet(fDropMonth) & ") is not part of this enrollment (" & _
                   FullSequenceText() & ")." & vbCrLf & "Choose one of the covered months.", vbExclamation
            Exit Function
        End If
        If SeqYearOffset(fDropPos) < 0 Then
            Notify MonthSheet(fDropMonth) & " belongs to LAST year's part of this enrollment." & vbCrLf & _
                   "Record that drop in last year's workbook, or choose a month of this year.", vbExclamation
            Exit Function
        End If
        If fDropPos = fMonths - 1 Then
            Notify MonthSheet(fDropMonth) & " is already the final month of this package, so the student " & _
                   "completes the package normally - no drop is needed." & vbCrLf & _
                   "Set Enrollment Status back to Active.", vbExclamation
            Exit Function
        End If
        fDropYear = SchoolYear() + SeqYearOffset(fDropPos)
    End If
    ReadForm = True
End Function

' Warns if the same student is already in this section in any target month
Private Function CheckDuplicates(months As Variant, excludeID As String) As Boolean
    Dim i As Long, r As Long, ws As Worksheet, sc As Long, ic As Long, hits As String
    sc = SecCol(fSection)
    ic = SecIdCol(fSection)
    For i = LBound(months) To UBound(months)
        Set ws = Worksheets(MonthSheet(months(i)))
        For r = FIRST_ROW To LAST_ROW
            If StrComp(CleanText(CStr(ws.Cells(r, sc + 1).Value)), fName, vbTextCompare) = 0 Then
                If excludeID = "" Or CStr(ws.Cells(r, ic).Value) <> excludeID Then
                    hits = hits & MonthSheet(months(i))
                    If CStr(ws.Cells(r, ic).Value) <> "" Then hits = hits & " (" & ws.Cells(r, ic).Value & ")"
                    hits = hits & ", "
                    Exit For
                End If
            End If
        Next r
    Next i
    If hits = "" Then
        CheckDuplicates = True
    Else
        hits = Left$(hits, Len(hits) - 2)
        CheckDuplicates = (Ask("POSSIBLE DUPLICATE" & vbCrLf & vbCrLf & fName & " is already in " & _
            SectionTitle(fSection) & " for: " & hits & "." & vbCrLf & vbCrLf & _
            "Add anyway? (Choose No unless this is a different student with the same name.)", _
            vbYesNo + vbExclamation + vbDefaultButton2) = vbYes)
    End If
End Function

' Every target month must have a free row BEFORE anything is written
Private Function CheckCapacity(months As Variant, excludeID As String, quiet As Boolean) As Boolean
    Dim i As Long, r As Long, ws As Worksheet, ic As Long, freeRows As Long, full As String
    ic = SecIdCol(fSection)
    For i = LBound(months) To UBound(months)
        Set ws = Worksheets(MonthSheet(months(i)))
        freeRows = 0
        For r = FIRST_ROW To LAST_ROW
            If RowIsEmpty(ws, r, fSection) Then
                freeRows = freeRows + 1
            ElseIf excludeID <> "" And CStr(ws.Cells(r, ic).Value) = excludeID Then
                freeRows = freeRows + 1
            End If
        Next r
        If freeRows = 0 Then full = full & MonthSheet(months(i)) & ", "
    Next i
    If full = "" Then
        CheckCapacity = True
    Else
        gLastMessage = "No free rows left in " & SectionTitle(fSection) & " for: " & Left$(full, Len(full) - 2) & "."
        If Not quiet Then Notify gLastMessage & vbCrLf & "Nothing was saved.", vbCritical
        CheckCapacity = False
    End If
End Function

'==============================================================================
'  MONTHLY SHEET OPERATIONS
'==============================================================================
Private Sub WriteToMonths(months As Variant, id As String)
    Dim i As Long, ws As Worksheet, r As Long, sc As Long
    sc = SecCol(fSection)
    For i = LBound(months) To UBound(months)
        Set ws = Worksheets(MonthSheet(months(i)))
        r = NextFreeRow(ws, fSection)
        If r = 0 Then Err.Raise vbObjectError + 513, , "No free row in " & ws.Name
        ws.Cells(r, sc + 1).Value = fName
        ws.Cells(r, sc + 2).Value = PackageLabel(fProgram, fPackage, fSection)
        If SecHasDiscount(fSection) Then
            ws.Cells(r, sc + 3).Value = fDiscount
            ws.Cells(r, sc + 4).Value = fNet
            ws.Cells(r, sc + 5).Value = fPerMonth
        Else
            ws.Cells(r, sc + 3).Value = fNet
            ws.Cells(r, sc + 4).Value = fPerMonth
        End If
        ws.Cells(r, SecIdCol(fSection)).Value = id
    Next i
End Sub

' Removes every row carrying this ID from all sections of all 12 months and
' moves the remaining rows up. Only values move; formulas/formatting stay.
Private Function RemoveFromMonths(id As String) As Long
    Dim m As Long, k As Long, ws As Worksheet, sc As Long, ic As Long, nCols As Long
    Dim data As Variant, keep() As Variant, r As Long, c As Long, outR As Long
    Dim removed As Long, nRows As Long, found As Boolean, kept() As Long

    nRows = LAST_ROW - FIRST_ROW + 1
    For m = 1 To 12
        Set ws = Worksheets(MonthSheet(m))
        For k = 1 To 3
            sc = SecCol(k)
            ic = SecIdCol(k)
            nCols = ic - sc                       ' Name .. ID
            data = ws.Range(ws.Cells(FIRST_ROW, sc + 1), ws.Cells(LAST_ROW, ic)).Value
            found = False
            For r = 1 To nRows
                If CStr(data(r, nCols)) = id Then found = True: Exit For
            Next r
            If found Then
                ReDim kept(1 To nRows)
                outR = 0
                For r = 1 To nRows
                    If CStr(data(r, nCols)) = id Then
                        removed = removed + 1
                    ElseIf Not RowArrayEmpty(data, r, nCols) Then
                        outR = outR + 1
                        kept(outR) = r
                    End If
                Next r
                ws.Range(ws.Cells(FIRST_ROW, sc + 1), ws.Cells(LAST_ROW, ic)).ClearContents
                If outR > 0 Then
                    ReDim keep(1 To outR, 1 To nCols)
                    For r = 1 To outR
                        For c = 1 To nCols
                            keep(r, c) = data(kept(r), c)
                        Next c
                    Next r
                    ws.Range(ws.Cells(FIRST_ROW, sc + 1), ws.Cells(FIRST_ROW + outR - 1, ic)).Value = keep
                End If
            End If
        Next k
    Next m
    RemoveFromMonths = removed
End Function

Private Function NextFreeRow(ws As Worksheet, k As Long) As Long
    Dim r As Long
    For r = FIRST_ROW To LAST_ROW
        If RowIsEmpty(ws, r, k) Then
            NextFreeRow = r
            Exit Function
        End If
    Next r
    NextFreeRow = 0
End Function

' A row is free only if EVERY data cell (Name..ID) is empty - never overwrite.
Private Function RowIsEmpty(ws As Worksheet, r As Long, k As Long) As Boolean
    Dim c As Long
    For c = SecCol(k) + 1 To SecIdCol(k)
        If CStr(ws.Cells(r, c).Value) <> "" Then
            RowIsEmpty = False
            Exit Function
        End If
    Next c
    RowIsEmpty = True
End Function

Private Function RowArrayEmpty(data As Variant, r As Long, nCols As Long) As Boolean
    Dim c As Long
    For c = 1 To nCols
        If CStr(data(r, c)) <> "" Then
            RowArrayEmpty = False
            Exit Function
        End If
    Next c
    RowArrayEmpty = True
End Function

'==============================================================================
'  ENROLLMENT HISTORY / STUDENT PROFILES
'==============================================================================
' IDs include the school year (UPK-2026-0001, STU-2026-0001) so they stay
' unique when carry-over rows move between yearly workbooks.
Private Function NextEnrollmentID() As String
    Dim pre As String
    pre = "UPK-" & SchoolYear() & "-"
    NextEnrollmentID = pre & Format$(MaxIdNumber(LC_EID, pre) + 1, "0000")
End Function

Private Function NextStudentID() As String
    Dim pre As String
    pre = "STU-" & SchoolYear() & "-"
    NextStudentID = pre & Format$(MaxIdNumber(LC_SID, pre) + 1, "0000")
End Function

Private Function MaxIdNumber(col As Long, prefix As String) As Long
    Dim lg As Worksheet, r As Long, s As String, n As Long, mx As Long
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, LC_EID).Value) = "" Then Exit For
        s = CStr(lg.Cells(r, col).Value)
        If UCase$(Left$(s, Len(prefix))) = prefix Then
            If IsNumeric(Mid$(s, Len(prefix) + 1)) Then
                n = CLng(Mid$(s, Len(prefix) + 1))
                If n > mx Then mx = n
            End If
        End If
    Next r
    MaxIdNumber = mx
End Function

Private Function NextLogRow() As Long
    Dim lg As Worksheet, r As Long
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, LC_EID).Value) = "" Then
            NextLogRow = r
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 514, , "Enrollment History is full."
End Function

Private Function FindLogRow(ByVal id As String) As Long
    Dim lg As Worksheet, r As Long, s As String
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        s = CStr(lg.Cells(r, LC_EID).Value)
        If s = "" Then Exit For
        If StrComp(s, id, vbTextCompare) = 0 Then
            FindLogRow = r
            Exit Function
        End If
    Next r
    FindLogRow = 0
End Function

' Student ID of an existing student with exactly this name ("" if new)
Private Function FindStudentID(ByVal nm As String) As String
    Dim lg As Worksheet, r As Long
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, LC_EID).Value) = "" Then Exit For
        If StrComp(CleanText(CStr(lg.Cells(r, LC_NAME).Value)), nm, vbTextCompare) = 0 Then
            If CStr(lg.Cells(r, LC_SID).Value) <> "" Then
                FindStudentID = CStr(lg.Cells(r, LC_SID).Value)
                Exit Function
            End If
        End If
    Next r
    FindStudentID = ""
End Function

' "STU-0001|STU-0004" - students whose name contains term, or whose Student ID
' or one of whose Enrollment IDs equals term
Private Function MatchingStudentIDs(ByVal term As String) As String
    Dim lg As Worksheet, r As Long, sid As String, res As String
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, LC_EID).Value) = "" Then Exit For
        sid = CStr(lg.Cells(r, LC_SID).Value)
        If sid <> "" Then
            If InStr(1, CStr(lg.Cells(r, LC_NAME).Value), term, vbTextCompare) > 0 _
               Or StrComp(sid, term, vbTextCompare) = 0 _
               Or StrComp(CStr(lg.Cells(r, LC_EID).Value), term, vbTextCompare) = 0 Then
                If InStr(1, "|" & res & "|", "|" & sid & "|", vbTextCompare) = 0 Then
                    If res <> "" Then res = res & "|"
                    res = res & sid
                End If
            End If
        End If
    Next r
    MatchingStudentIDs = res
End Function

Private Function StudentName(ByVal sid As String) As String
    Dim r As Long
    r = LastEnrollmentRow(sid)
    If r > 0 Then StudentName = CStr(Worksheets(SH_LOG).Cells(r, LC_NAME).Value)
End Function

' Latest non-cancelled enrollment row of a student (latest of any if all cancelled)
Private Function LastEnrollmentRow(ByVal sid As String) As Long
    Dim lg As Worksheet, r As Long, anyRow As Long, best As Long
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, LC_EID).Value) = "" Then Exit For
        If StrComp(CStr(lg.Cells(r, LC_SID).Value), sid, vbTextCompare) = 0 Then
            anyRow = r
            If CStr(lg.Cells(r, LC_STATUS).Value) <> "Cancelled" Then best = r
        End If
    Next r
    If best = 0 Then best = anyRow
    LastEnrollmentRow = best
End Function

Private Function LastEnrollmentText(ByVal sid As String) As String
    Dim r As Long, lg As Worksheet, cov As String
    r = LastEnrollmentRow(sid)
    If r = 0 Then Exit Function
    Set lg = Worksheets(SH_LOG)
    cov = CoveredText(r)
    LastEnrollmentText = lg.Cells(r, LC_PROGRAM).Value & " " & lg.Cells(r, LC_PACKAGE).Value & ", " & _
        IIf(cov = "", "(no months)", MonthRangeText(Replace(cov, " (next year)", ""))) & _
        " [" & lg.Cells(r, LC_STATUS).Value & "]"
End Function

' Month after the last month the student attended in that enrollment
Private Function NextMonthAfter(r As Long) As Long
    Dim cov As String, parts As Variant, m As Long, lg As Worksheet
    Set lg = Worksheets(SH_LOG)
    cov = CStr(lg.Cells(r, LC_NEXTYEAR).Value)
    If cov = "" Then cov = CStr(lg.Cells(r, LC_COVERED).Value)
    If cov <> "" Then
        parts = Split(cov, ",")
        m = MonthIndex(Trim$(CStr(parts(UBound(parts)))))
    Else
        m = MonthIndex(CStr(lg.Cells(r, LC_START).Value)) - 1
        If m = 0 Then m = 12
    End If
    If m = 0 Then
        NextMonthAfter = 0
    Else
        NextMonthAfter = (m Mod 12) + 1
    End If
End Function

Private Sub LogNew(ByVal id As String, ByVal sid As String, ByVal kind As String, months As Variant, _
                   nextM As Variant, ByVal submitted As Variant)
    Dim lg As Worksheet, r As Long
    Set lg = Worksheets(SH_LOG)
    r = NextLogRow()
    lg.Cells(r, LC_EID).Value = id
    lg.Cells(r, LC_SID).Value = sid
    lg.Cells(r, LC_DATE).Value = submitted
    lg.Cells(r, LC_TYPE).Value = kind
    WriteLogDetails r, months, nextM
    Select Case kind
        Case "New": lg.Cells(r, LC_REMARKS).Value = "Submitted."
        Case "Re-enrollment": lg.Cells(r, LC_REMARKS).Value = "Re-enrolled."
        Case Else: lg.Cells(r, LC_REMARKS).Value = "Carried over from " & (SchoolYear() + fStartOffset) & " on " & Format$(Now, "yyyy-mm-dd") & "."
    End Select
End Sub

Private Sub LogUpdate(r As Long, months As Variant, nextM As Variant)
    Dim lg As Worksheet
    Set lg = Worksheets(SH_LOG)
    WriteLogDetails r, months, nextM
    If fStatus = "Dropped" Then
        lg.Cells(r, LC_REMARKS).Value = Trim$(lg.Cells(r, LC_REMARKS).Value & " Dropped - last month " & _
            MonthSheet(fDropMonth) & " " & fDropYear & " (" & Format$(Now, "yyyy-mm-dd") & ").")
    Else
        lg.Cells(r, LC_REMARKS).Value = Trim$(lg.Cells(r, LC_REMARKS).Value & " Edited " & Format$(Now, "yyyy-mm-dd") & ".")
    End If
End Sub

Private Sub WriteLogDetails(r As Long, months As Variant, nextM As Variant)
    Dim lg As Worksheet
    Set lg = Worksheets(SH_LOG)
    lg.Cells(r, LC_PROGRAM).Value = fProgram
    lg.Cells(r, LC_SECTION).Value = SectionTitle(fSection)
    lg.Cells(r, LC_NAME).Value = fName
    lg.Cells(r, LC_PACKAGE).Value = fPackage
    lg.Cells(r, LC_MONTHS).Value = fMonths
    lg.Cells(r, LC_START).Value = MonthSheet(fStartMonth)
    lg.Cells(r, LC_STARTYEAR).Value = SchoolYear() + fStartOffset
    lg.Cells(r, LC_COVERED).Value = IIf(ArrLen(months) = 0, "", MonthListText(months))
    lg.Cells(r, LC_NEXTYEAR).Value = IIf(ArrLen(nextM) = 0, "", MonthListText(nextM))
    lg.Cells(r, LC_FEE).Value = fFee
    lg.Cells(r, LC_DISCOUNT).Value = fDiscount
    lg.Cells(r, LC_NET).Value = fNet
    lg.Cells(r, LC_PERMONTH).Value = fPerMonth
    lg.Cells(r, LC_STATUS).Value = fStatus
    lg.Cells(r, LC_DROP).Value = IIf(fDropMonth > 0, MonthSheet(fDropMonth), "")
    lg.Cells(r, LC_DROPYEAR).Value = IIf(fDropMonth > 0, fDropYear, "")
    lg.Cells(r, LC_UPDATED).Value = Now
End Sub

'==============================================================================
'  YEAR-END CARRY-OVER
'==============================================================================
' Rebuilds the End of Year sheet from Enrollment History: every enrollment
' (not cancelled) that still has months in NEXT year.
Public Sub RefreshEndOfYear()
    Dim eo As Worksheet, lg As Worksheet, r As Long, outR As Long
    On Error GoTo Done
    Set eo = Worksheets(SH_EOY)
    Set lg = Worksheets(SH_LOG)
    eo.Range(eo.Cells(CARRY_FIRST, 1), eo.Cells(CARRY_LAST, 17)).ClearContents
    outR = CARRY_FIRST
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, LC_EID).Value) = "" Then Exit For
        If CStr(lg.Cells(r, LC_STATUS).Value) <> "Cancelled" And CStr(lg.Cells(r, LC_NEXTYEAR).Value) <> "" _
           And outR <= CARRY_LAST Then
            eo.Cells(outR, 1).Value = lg.Cells(r, LC_EID).Value
            eo.Cells(outR, 2).Value = lg.Cells(r, LC_SID).Value
            eo.Cells(outR, 3).Value = lg.Cells(r, LC_NAME).Value
            eo.Cells(outR, 4).Value = lg.Cells(r, LC_PROGRAM).Value
            eo.Cells(outR, 5).Value = lg.Cells(r, LC_PACKAGE).Value
            eo.Cells(outR, 6).Value = lg.Cells(r, LC_MONTHS).Value
            eo.Cells(outR, 7).Value = lg.Cells(r, LC_START).Value
            eo.Cells(outR, 8).Value = lg.Cells(r, LC_STARTYEAR).Value
            eo.Cells(outR, 9).Value = lg.Cells(r, LC_FEE).Value
            eo.Cells(outR, 10).Value = lg.Cells(r, LC_DISCOUNT).Value
            eo.Cells(outR, 11).Value = lg.Cells(r, LC_NET).Value
            eo.Cells(outR, 12).Value = lg.Cells(r, LC_PERMONTH).Value
            eo.Cells(outR, 13).Value = lg.Cells(r, LC_NEXTYEAR).Value
            eo.Cells(outR, 14).Value = lg.Cells(r, LC_STATUS).Value
            eo.Cells(outR, 15).Value = lg.Cells(r, LC_DROP).Value
            eo.Cells(outR, 16).Value = lg.Cells(r, LC_DROPYEAR).Value
            eo.Cells(outR, 17).Value = lg.Cells(r, LC_DATE).Value
            outR = outR + 1
        End If
    Next r
Done:
End Sub

' Next year's workbook: adds the pasted carry-over rows to this year's
' monthly sheets and Enrollment History. Safe to run more than once.
Public Sub ImportCarryOver()
    Dim bo As Worksheet, r As Long, n As Long, nOk As Long, nSkip As Long, nErr As Long
    Dim yrs As String, yr As Variant, eid As String, sid As String, existing As String
    Dim months As Variant, i As Long, res As String, k As Long

    Set bo = Worksheets(SH_BOY)
    For r = CARRY_FIRST To CARRY_LAST
        If Trim$(CStr(bo.Cells(r, 1).Value)) <> "" Then
            n = n + 1
            yr = bo.Cells(r, 8).Value
            If IsNumeric(yr) And CStr(yr) <> "" Then
                If CLng(yr) <> SchoolYear() - 1 And InStr(yrs, "|" & yr & "|") = 0 Then yrs = yrs & "|" & yr & "|"
            End If
        End If
    Next r
    If n = 0 Then
        Notify "Nothing to import." & vbCrLf & vbCrLf & _
               "Paste LAST year's End of Year rows at cell A9 of this sheet, then click IMPORT CARRY-OVER.", vbExclamation
        Exit Sub
    End If
    If yrs <> "" Then
        If Ask("Some rows started in " & Replace(Replace(Mid$(yrs, 2, Len(yrs) - 2), "||", ", "), "|", "") & _
               ", but this workbook's School Year is " & SchoolYear() & " (Settings!B2)." & vbCrLf & _
               "Carry-over rows normally come from " & (SchoolYear() - 1) & "." & vbCrLf & vbCrLf & _
               "Import anyway?", vbYesNo + vbExclamation + vbDefaultButton2) <> vbYes Then Exit Sub
    End If
    If Ask("Import " & n & " carry-over row(s) into " & SchoolYear() & "?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    For r = CARRY_FIRST To CARRY_LAST
        eid = UCase$(Trim$(CStr(bo.Cells(r, 1).Value)))
        If eid <> "" And Left$(CStr(bo.Cells(r, 18).Value), 8) = "Imported" Then
            nSkip = nSkip + 1                       ' imported earlier - keep its note
        ElseIf eid <> "" Then
            res = ""
            If FindLogRow(eid) > 0 Then
                res = "Skipped - already in Enrollment History"
                nSkip = nSkip + 1
            Else
                fName = CleanText(CStr(bo.Cells(r, 3).Value))
                fProgram = CleanText(CStr(bo.Cells(r, 4).Value))
                fPackage = CleanText(CStr(bo.Cells(r, 5).Value))
                fMonths = CLng(Val(CStr(bo.Cells(r, 6).Value)))
                fStartMonth = MonthIndex(CStr(bo.Cells(r, 7).Value))
                fStartOffset = CLng(Val(CStr(bo.Cells(r, 8).Value))) - SchoolYear()
                fFee = Val(CStr(bo.Cells(r, 9).Value))
                fDiscount = Val(CStr(bo.Cells(r, 10).Value))
                fNet = Val(CStr(bo.Cells(r, 11).Value))
                fPerMonth = Val(CStr(bo.Cells(r, 12).Value))
                fStatus = IIf(CStr(bo.Cells(r, 14).Value) = "Dropped", "Dropped", "Active")
                fDropMonth = 0: fDropPos = -1: fDropYear = 0
                fSection = ProgramSection(fProgram)
                If fStatus = "Dropped" Then
                    fDropMonth = MonthIndex(CStr(bo.Cells(r, 15).Value))
                    fDropYear = CLng(Val(CStr(bo.Cells(r, 16).Value)))
                    For i = 0 To fMonths - 1
                        If SeqMonth(i) = fDropMonth And SchoolYear() + SeqYearOffset(i) = fDropYear Then fDropPos = i
                    Next i
                End If
                If fSection = 0 Then
                    res = "ERROR - program '" & fProgram & "' is not in the Settings Programs table"
                ElseIf fName = "" Or fStartMonth = 0 Or fMonths < 1 Or fMonths > 12 Then
                    res = "ERROR - row is incomplete (name, start month or months)"
                Else
                    months = ThisYearList(AttendedCount())
                    If ArrLen(months) = 0 Then
                        res = "Skipped - no months fall in " & SchoolYear()
                    ElseIf Not CheckCapacity(months, "", True) Then
                        res = "ERROR - " & gLastMessage
                    ElseIf DuplicateMonths(months) <> "" Then
                        res = "Skipped - " & fName & " is already on the sheet for " & DuplicateMonths(months)
                    End If
                End If
                If res <> "" Then
                    If Left$(res, 5) = "ERROR" Then nErr = nErr + 1 Else nSkip = nSkip + 1
                Else
                    existing = FindStudentID(fName)
                    sid = UCase$(Trim$(CStr(bo.Cells(r, 2).Value)))
                    If existing <> "" Then
                        sid = existing
                    ElseIf sid = "" Or StudentIdUsed(sid) Then
                        sid = NextStudentID()
                    End If
                    WriteToMonths months, eid
                    LogNew eid, sid, "Carry-over", months, Array(), IIf(IsDate(bo.Cells(r, 17).Value), bo.Cells(r, 17).Value, Now)
                    res = "Imported " & Format$(Now, "yyyy-mm-dd") & " - " & MonthListText(months)
                    nOk = nOk + 1
                End If
            End If
            bo.Cells(r, 18).Value = res
        End If
    Next r
    Application.ScreenUpdating = True
    Notify "Carry-over import finished." & vbCrLf & vbCrLf & _
           "Imported: " & nOk & vbCrLf & "Skipped: " & nSkip & vbCrLf & "Errors: " & nErr & vbCrLf & vbCrLf & _
           "See the Import Result column for details.", IIf(nErr > 0, vbExclamation, vbInformation)
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify "Import stopped because of an error: " & Err.Description, vbCritical
End Sub

' Months (of the list) where fName is already in fSection; "" if none
Private Function DuplicateMonths(months As Variant) As String
    Dim i As Long, r As Long, ws As Worksheet, sc As Long, s As String
    sc = SecCol(fSection)
    For i = LBound(months) To UBound(months)
        Set ws = Worksheets(MonthSheet(months(i)))
        For r = FIRST_ROW To LAST_ROW
            If StrComp(CleanText(CStr(ws.Cells(r, sc + 1).Value)), fName, vbTextCompare) = 0 Then
                If s <> "" Then s = s & ", "
                s = s & MonthSheet(months(i))
                Exit For
            End If
        Next r
    Next i
    DuplicateMonths = s
End Function

Private Function StudentIdUsed(ByVal sid As String) As Boolean
    Dim lg As Worksheet, r As Long
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, LC_EID).Value) = "" Then Exit For
        If StrComp(CStr(lg.Cells(r, LC_SID).Value), sid, vbTextCompare) = 0 Then
            StudentIdUsed = True
            Exit Function
        End If
    Next r
End Function

'==============================================================================
'  MONTH SEQUENCE (position i = 0 .. fMonths-1 of the package)
'==============================================================================
Private Function SchoolYear() As Long
    Dim v As Variant
    v = SafeGet("cfgSchoolYear")
    If IsNumeric(v) And CStr(v) <> "" Then SchoolYear = CLng(v) Else SchoolYear = Year(Date)
End Function

Private Function SeqPos(ByVal i As Long) As Long
    SeqPos = (fStartMonth - 1) + 12 * fStartOffset + i
End Function

Private Function SeqMonth(ByVal i As Long) As Long
    SeqMonth = (((SeqPos(i) Mod 12) + 12) Mod 12) + 1
End Function

' -1 = last year, 0 = this workbook's year, 1 = next year
Private Function SeqYearOffset(ByVal i As Long) As Long
    Dim p As Long
    p = SeqPos(i)
    If p < 0 Then
        SeqYearOffset = -((-p - 1) \ 12) - 1
    Else
        SeqYearOffset = p \ 12
    End If
End Function

' Number of package months actually attended (all, or up to the last month)
Private Function AttendedCount() As Long
    If fDropPos >= 0 Then AttendedCount = fDropPos + 1 Else AttendedCount = fMonths
End Function

Private Function ThisYearList(ByVal n As Long) As Variant
    ThisYearList = YearList(n, 0)
End Function

Private Function NextYearList(ByVal n As Long) As Variant
    NextYearList = YearList(n, 1)
End Function

Private Function YearList(ByVal n As Long, ByVal yearOffset As Long) As Variant
    Dim a() As Long, c As Long, i As Long
    For i = 0 To n - 1
        If SeqYearOffset(i) = yearOffset Then c = c + 1
    Next i
    If c = 0 Then
        YearList = Array()
        Exit Function
    End If
    ReDim a(1 To c)
    c = 0
    For i = 0 To n - 1
        If SeqYearOffset(i) = yearOffset Then
            c = c + 1
            a(c) = SeqMonth(i)
        End If
    Next i
    YearList = a
End Function

Private Function FullSequenceText() As String
    Dim i As Long, s As String
    For i = 0 To fMonths - 1
        If s <> "" Then s = s & ", "
        s = s & MonthSheet(SeqMonth(i))
    Next i
    FullSequenceText = s
End Function

Private Function ListOrNone(months As Variant) As String
    If ArrLen(months) = 0 Then ListOrNone = "(no months)" Else ListOrNone = MonthListText(months)
End Function

' This-year months plus next-year months of a history row
Private Function CoveredText(ByVal r As Long) As String
    Dim lg As Worksheet, a As String, b As String
    Set lg = Worksheets(SH_LOG)
    a = CStr(lg.Cells(r, LC_COVERED).Value)
    b = CStr(lg.Cells(r, LC_NEXTYEAR).Value)
    If b <> "" Then
        b = Replace(b, ",", " (next year),") & " (next year)"
        If a <> "" Then CoveredText = a & ", " & b Else CoveredText = b
    Else
        CoveredText = a
    End If
End Function

'==============================================================================
'  SETTINGS LOOKUPS
'==============================================================================
' Section (1..3) of a program from the Settings Programs table; 0 if unknown
Private Function ProgramSection(ByVal prg As String) As Long
    Dim t As Variant, i As Long
    ProgramSection = 0
    If prg = "" Then Exit Function
    t = Range("ProgTable").Value
    For i = 1 To UBound(t, 1)
        If StrComp(CleanText(CStr(t(i, 1))), prg, vbTextCompare) = 0 Then
            ProgramSection = SectionIndex(CStr(t(i, 2)))
            Exit Function
        End If
    Next i
End Function

Private Function LookupPackage(ByVal prg As String, ByVal pkg As String, months As Long, fee As Double) As Boolean
    Dim progs As Variant, names As Variant, mos As Variant, fees As Variant, i As Long
    progs = Range("PkgProgram").Value
    names = Range("PkgName").Value
    mos = Range("PkgMonths").Value
    fees = Range("PkgFee").Value
    For i = 1 To UBound(progs, 1)
        If StrComp(CleanText(CStr(progs(i, 1))), prg, vbTextCompare) = 0 And _
           StrComp(CleanText(CStr(names(i, 1))), pkg, vbTextCompare) = 0 Then
            If IsNumeric(mos(i, 1)) And CStr(mos(i, 1)) <> "" Then months = CLng(mos(i, 1)) Else months = 0
            If IsNumeric(fees(i, 1)) And CStr(fees(i, 1)) <> "" Then fee = CDbl(fees(i, 1)) Else fee = 0
            LookupPackage = True
            Exit Function
        End If
    Next i
    LookupPackage = False
End Function

' Label shown in the Package column: "6 Months", "VIP 6 Months", "4x 1 Month"
Private Function PackageLabel(ByVal prg As String, ByVal pkg As String, ByVal k As Long) As String
    Dim vpart As String
    vpart = Trim$(Replace(prg, SecBaseProgram(k), "", 1, -1, vbTextCompare))
    If vpart = prg Then vpart = ""
    If vpart = "" Then PackageLabel = pkg Else PackageLabel = vpart & " " & pkg
End Function

'==============================================================================
'  HELPERS
Private Function ArrLen(a As Variant) As Long
    On Error GoTo NoItems
    ArrLen = UBound(a) - LBound(a) + 1
    Exit Function
NoItems:
    ArrLen = 0
End Function

Private Function MonthSheet(ByVal m As Long) As String
    If m < 1 Or m > 12 Then Exit Function
    MonthSheet = Choose(m, "January", "February", "March", "April", "May", "June", _
                        "July", "August", "September", "October", "November", "December")
End Function

Private Function MonthIndex(ByVal s As String) As Long
    Dim m As Long
    For m = 1 To 12
        If StrComp(Trim$(s), MonthSheet(m), vbTextCompare) = 0 Then
            MonthIndex = m
            Exit Function
        End If
    Next m
    MonthIndex = 0
End Function

Private Function MonthListText(months As Variant) As String
    Dim i As Long, s As String
    For i = LBound(months) To UBound(months)
        If s <> "" Then s = s & ", "
        s = s & MonthSheet(months(i))
    Next i
    MonthListText = s
End Function

' "January, February, March" -> "January-March"
Private Function MonthRangeText(ByVal cov As String) As String
    Dim parts As Variant
    parts = Split(cov, ",")
    If UBound(parts) = 0 Then
        MonthRangeText = Trim$(CStr(parts(0)))
    Else
        MonthRangeText = Trim$(CStr(parts(0))) & "-" & Trim$(CStr(parts(UBound(parts))))
    End If
End Function

Private Function SecCol(ByVal k As Long) As Long
    SecCol = Choose(k, COL_TODDLER, COL_PRESCHOOL, COL_TUTOR)
End Function

Private Function SecHasDiscount(ByVal k As Long) As Boolean
    SecHasDiscount = (k <> 3)
End Function

' Hidden Enrollment-ID column = the column right after "Per Month"
Private Function SecIdCol(ByVal k As Long) As Long
    If SecHasDiscount(k) Then SecIdCol = SecCol(k) + 6 Else SecIdCol = SecCol(k) + 5
End Function

Private Function SectionTitle(ByVal k As Long) As String
    If k < 1 Or k > 3 Then Exit Function
    SectionTitle = Choose(k, "Toddlers Group", "Preschool Group", "One-on-One Tutoring")
End Function

Private Function SecBaseProgram(ByVal k As Long) As String
    SecBaseProgram = Choose(k, "Toddler PlayTrack", "Preschool PlayTrack", "One-on-One Tutoring")
End Function

Private Function SectionIndex(ByVal s As String) As Long
    Dim k As Long
    s = LCase$(Trim$(s))
    For k = 1 To 3
        If s = LCase$(SectionTitle(k)) Then
            SectionIndex = k
            Exit Function
        End If
    Next k
    ' tolerant matches
    If InStr(s, "toddler") > 0 Then
        SectionIndex = 1
    ElseIf InStr(s, "preschool") > 0 Then
        SectionIndex = 2
    ElseIf InStr(s, "one") > 0 Or InStr(s, "tutor") > 0 Then
        SectionIndex = 3
    End If
End Function

Private Function FormSummary(months As Variant, nextM As Variant) As String
    FormSummary = "Program: " & fProgram & "  (" & SectionTitle(fSection) & ")" & vbCrLf & _
                  "Full Name: " & fName & vbCrLf & _
                  "Package: " & fPackage & " = " & fMonths & " month(s)" & vbCrLf & _
                  "This year: " & ListOrNone(months) & vbCrLf & _
                  IIf(ArrLen(nextM) > 0, "Next year (End of Year sheet): " & ListOrNone(nextM) & vbCrLf, "") & _
                  "Program Fee: " & Peso(fFee) & "   Discount: " & Peso(fDiscount) & vbCrLf & _
                  "Net Price: " & Peso(fNet) & "   Per month: " & Peso(fPerMonth) & vbCrLf
End Function

Private Function Peso(ByVal x As Double) As String
    Peso = ChrW(8369) & Format$(x, "#,##0.00")
End Function

Private Function CleanText(ByVal s As String) As String
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = s
End Function

' Never fails, even if a named cell was deleted - this is what keeps
' CLEAR FORM working.
Private Sub ClearInputs()
    SafeSet "inProgram", ""
    SafeSet "inName", ""
    SafeSet "inPackage", ""
    SafeSet "inStartMonth", ""
    SafeSet "inDiscount", ""
    SafeSet "inStatus", ""
    SafeSet "inDropMonth", ""
End Sub

Private Sub SafeSet(ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    Range(nm).Value = v
End Sub

Private Function SafeGet(ByVal nm As String) As Variant
    On Error Resume Next
    SafeGet = ""
    SafeGet = Range(nm).Value
    If IsError(SafeGet) Then SafeGet = ""
End Function

Private Function Ask(msg As String, buttons As Long) As VbMsgBoxResult
    If gSilent Then
        gLastMessage = msg
        Ask = gSilentAnswer
    Else
        Ask = MsgBox(msg, buttons, APP_TITLE)
    End If
End Function

Private Sub Notify(msg As String, icon As Long)
    gLastMessage = msg
    If Not gSilent Then MsgBox msg, icon, APP_TITLE
End Sub

Private Sub DeleteButtons(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "btn" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Function FindShape(ws As Worksheet, shpName As String) As Object
    Dim i As Long
    Set FindShape = Nothing
    For i = 1 To ws.Shapes.Count
        If ws.Shapes(i).Name = shpName Then
            Set FindShape = ws.Shapes(i)
            Exit Function
        End If
    Next i
End Function

' Shapes are late-bound (As Object) so no extra library reference is needed.
Private Sub PlaceButton(ws As Worksheet, target As Range, shpName As String, caption As String, _
                        macroName As String, fillColor As Long, fontSize As Single)
    Dim shp As Object, tf As Object
    Set shp = FindShape(ws, shpName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(5, target.Left + 3, target.Top + 2, target.Width - 6, target.Height - 4) ' 5 = rounded rectangle
        shp.Name = shpName
        shp.Fill.ForeColor.RGB = fillColor
        shp.Line.Visible = 0
        shp.Placement = 2
        Set tf = shp.TextFrame2
        tf.VerticalAnchor = 3
        tf.TextRange.Text = caption
        tf.TextRange.ParagraphFormat.Alignment = 2
        tf.TextRange.Font.Bold = -1
        tf.TextRange.Font.Size = fontSize
        tf.TextRange.Font.Name = "Arial"
        tf.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End If
    shp.OnAction = macroName
End Sub
