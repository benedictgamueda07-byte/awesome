Attribute VB_Name = "UpKidzEnrollment"
'==============================================================================
'  UPKIDZ PHILIPPINES - ENROLLMENT MANAGEMENT SYSTEM
'  VBA module for UpKidz_Enrollment_System workbook
'
'  Public macros (assign to buttons):
'    InstallUpKidzSystem  - run ONCE after importing: creates all buttons
'    SubmitEnrollment     - add a new enrollment to every covered month
'    ClearForm            - clear the Dashboard form / leave edit mode
'    LoadEnrollmentForEdit- load an Enrollment ID into the form
'    UpdateEnrollment     - apply edits to every affected month
'    CancelEnrollment     - remove an enrollment from every month
'    SearchStudent        - find a student across all 12 monthly sheets
'    GoToDashboard        - navigation helper
'
'  Every record on a monthly sheet carries its Enrollment ID in a hidden
'  column, so Edit and Cancel always touch exactly the right rows in every
'  month and never leave stale or duplicate copies behind.
'==============================================================================
Option Explicit

'---- layout (must match build_workbook.py) -----------------------------------
Private Const FIRST_ROW As Long = 7          ' first enrollment row on month sheets
Private Const LAST_ROW As Long = 106         ' last enrollment row (100 slots)
Private Const COL_TODDLER As Long = 2        ' column B  = "No." of Toddlers Group
Private Const COL_PRESCHOOL As Long = 10     ' column J  = "No." of Preschool Group
Private Const COL_TUTOR As Long = 18         ' column R  = "No." of One-on-One

Private Const SH_DASH As String = "Dashboard"
Private Const SH_LOG As String = "Enrollment History"
Private Const SH_SEARCH As String = "Student Search"
Private Const LOG_FIRST_ROW As Long = 5
Private Const LOG_LAST_ROW As Long = 1005
Private Const SEARCH_FIRST_ROW As Long = 6

Private Const PRG_TODDLER As String = "Toddler PlayTrack"
Private Const PRG_PRESCHOOL As String = "Preschool PlayTrack"
Private Const PRG_TUTOR As String = "One-on-One Tutoring"

Private Const APP_TITLE As String = "UpKidz Enrollment System"

'---- current form data (filled by ReadForm) ----------------------------------
Private fProgram As String
Private fName As String
Private fSchedule As String
Private fPackage As String
Private fStartMonth As Long
Private fDuration As Long
Private fFee As Double
Private fDiscount As Double
Private fNet As Double

'---- test hook: when True, prompts are auto-answered and messages are stored -
Public gSilent As Boolean
Public gSilentAnswer As VbMsgBoxResult
Public gLastMessage As String

'==============================================================================
'  PUBLIC MACROS
'==============================================================================
Public Sub SubmitEnrollment()
    Dim months As Variant, newID As String, msg As String

    If Len(CStr(Range("inEditID").Value)) > 0 Then
        Notify "You are currently editing enrollment " & Range("inEditID").Value & "." & vbCrLf & _
               "Click UPDATE ENROLLMENT to save the changes, or CLEAR FORM to start a new enrollment.", vbExclamation
        Exit Sub
    End If
    If Not ReadForm() Then Exit Sub
    months = CoveredMonths(fStartMonth, fDuration)

    If Not CheckDuplicates(months, "") Then Exit Sub
    If Not CheckCapacity(months, "") Then Exit Sub

    msg = "Please confirm this enrollment:" & vbCrLf & vbCrLf & FormSummary(months) & vbCrLf & _
          "Submit now?"
    If Ask(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    newID = NextEnrollmentID()
    WriteToMonths months, newID
    LogNew newID, months
    Application.ScreenUpdating = True

    ClearInputs
    Notify "Enrollment submitted successfully!" & vbCrLf & vbCrLf & _
           "Enrollment ID: " & newID & vbCrLf & _
           "Student: " & fName & vbCrLf & _
           "Section: " & SectionTitle(fProgram) & vbCrLf & _
           "Added to: " & MonthListText(months), vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify "Submission stopped because of an error: " & Err.Description, vbCritical
End Sub

Public Sub ClearForm()
    ClearInputs
    Range("inEditID").Value = ""
End Sub

Public Sub LoadEnrollmentForEdit()
    Dim id As String, r As Long, lg As Worksheet
    id = UCase$(Trim$(CStr(Range("inManageID").Value)))
    If id = "" Then
        Notify "Type the Enrollment ID (e.g. UPK-0001) in the Enrollment ID box first.", vbExclamation
        Exit Sub
    End If
    r = FindLogRow(id)
    If r = 0 Then
        Notify "Enrollment ID " & id & " was not found in Enrollment History.", vbExclamation
        Exit Sub
    End If
    Set lg = Worksheets(SH_LOG)
    If lg.Cells(r, 13).Value = "Cancelled" Then
        Notify "Enrollment " & id & " is cancelled and cannot be edited." & vbCrLf & _
               "Submit it again as a new enrollment if needed.", vbExclamation
        Exit Sub
    End If
    Range("inProgram").Value = lg.Cells(r, 3).Value
    Range("inName").Value = lg.Cells(r, 4).Value
    Range("inSchedule").Value = lg.Cells(r, 5).Value
    Range("inPackage").Value = lg.Cells(r, 6).Value
    Range("inStartMonth").Value = lg.Cells(r, 7).Value
    Range("inDuration").Value = lg.Cells(r, 8).Value
    Range("inFee").Value = lg.Cells(r, 10).Value
    Range("inDiscount").Value = lg.Cells(r, 11).Value
    Range("inEditID").Value = id
    Range("inManageID").Value = id
    Notify "Enrollment " & id & " loaded into the form." & vbCrLf & _
           "Change any field, then click UPDATE ENROLLMENT.", vbInformation
End Sub

Public Sub UpdateEnrollment()
    Dim id As String, r As Long, months As Variant, removed As Long, msg As String
    id = CStr(Range("inEditID").Value)
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
    If Not ReadForm() Then Exit Sub
    months = CoveredMonths(fStartMonth, fDuration)
    If Not CheckDuplicates(months, id) Then Exit Sub
    If Not CheckCapacity(months, id) Then Exit Sub

    msg = "Update enrollment " & id & " with these details?" & vbCrLf & vbCrLf & FormSummary(months) & vbCrLf & _
          "The old entries will be replaced in every month."
    If Ask(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    removed = RemoveFromMonths(id)
    WriteToMonths months, id
    LogUpdate r, months
    Application.ScreenUpdating = True

    ClearInputs
    Range("inEditID").Value = ""
    Notify "Enrollment " & id & " updated." & vbCrLf & _
           "Old entries removed: " & removed & vbCrLf & _
           "Now recorded in: " & MonthListText(months), vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify "Update stopped because of an error: " & Err.Description, vbCritical
End Sub

Public Sub CancelEnrollment()
    Dim id As String, r As Long, lg As Worksheet, removed As Long
    id = UCase$(Trim$(CStr(Range("inManageID").Value)))
    If id = "" Then id = CStr(Range("inEditID").Value)
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
    If lg.Cells(r, 13).Value = "Cancelled" Then
        Notify "Enrollment " & id & " is already cancelled.", vbInformation
        Exit Sub
    End If
    If Ask("Cancel this enrollment?" & vbCrLf & vbCrLf & _
           "ID: " & id & vbCrLf & _
           "Student: " & lg.Cells(r, 4).Value & vbCrLf & _
           "Program: " & lg.Cells(r, 3).Value & vbCrLf & _
           "Months: " & lg.Cells(r, 9).Value & vbCrLf & vbCrLf & _
           "The student will be removed from every month listed. This cannot be undone.", _
           vbYesNo + vbExclamation + vbDefaultButton2) <> vbYes Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    removed = RemoveFromMonths(id)
    lg.Cells(r, 13).Value = "Cancelled"
    lg.Cells(r, 14).Value = Now
    lg.Cells(r, 15).Value = Trim$(lg.Cells(r, 15).Value & " Cancelled " & Format$(Now, "yyyy-mm-dd") & ".")
    Application.ScreenUpdating = True

    If CStr(Range("inEditID").Value) = id Then
        ClearInputs
        Range("inEditID").Value = ""
    End If
    Range("inManageID").Value = ""
    Notify "Enrollment " & id & " cancelled." & vbCrLf & _
           "Entries removed from monthly sheets: " & removed, vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Notify "Cancellation stopped because of an error: " & Err.Description, vbCritical
End Sub

Public Sub SearchStudent()
    Dim term As String, ss As Worksheet, m As Long, k As Long, r As Long
    Dim ws As Worksheet, sc As Long, prg As String, nm As String, outR As Long
    Dim id As String, hasDisc As Boolean, lastSearchRow As Long

    term = Trim$(CStr(Range("inSearch").Value))
    If term = "" Then
        Notify "Type all or part of a student's name in the Search Student box.", vbExclamation
        Exit Sub
    End If
    Set ss = Worksheets(SH_SEARCH)
    lastSearchRow = SEARCH_FIRST_ROW + 499
    ss.Range(ss.Cells(SEARCH_FIRST_ROW, 1), ss.Cells(lastSearchRow, 10)).ClearContents
    ss.Range("B3").Value = term
    outR = SEARCH_FIRST_ROW

    For m = 1 To 12
        Set ws = Worksheets(MonthSheet(m))
        For k = 1 To 3
            prg = ProgramByIndex(k)
            sc = SectionCol(prg)
            hasDisc = SectionHasDiscount(prg)
            For r = FIRST_ROW To LAST_ROW
                nm = CStr(ws.Cells(r, sc + 1).Value)
                If nm <> "" Then
                    If InStr(1, nm, term, vbTextCompare) > 0 Then
                        If outR > lastSearchRow Then GoTo Done
                        ss.Cells(outR, 1).Value = MonthSheet(m)
                        ss.Cells(outR, 2).Value = SectionTitle(prg)
                        ss.Cells(outR, 3).Value = ws.Cells(r, sc).Value
                        ss.Cells(outR, 4).Value = nm
                        ss.Cells(outR, 5).Value = ws.Cells(r, sc + 2).Value
                        ss.Cells(outR, 6).Value = ws.Cells(r, sc + 3).Value
                        If hasDisc Then
                            ss.Cells(outR, 7).Value = ws.Cells(r, sc + 4).Value
                            ss.Cells(outR, 8).Value = ws.Cells(r, sc + 5).Value
                        Else
                            ss.Cells(outR, 7).Value = ""
                            ss.Cells(outR, 8).Value = ws.Cells(r, sc + 4).Value
                        End If
                        id = CStr(ws.Cells(r, IdCol(prg)).Value)
                        ss.Cells(outR, 9).Value = id
                        If id <> "" Then
                            ss.Cells(outR, 10).Value = LogStatus(id)
                        Else
                            ss.Cells(outR, 10).Value = "Manual entry"
                        End If
                        outR = outR + 1
                    End If
                End If
            Next r
        Next k
    Next m
Done:
    If outR = SEARCH_FIRST_ROW Then
        ss.Range("D3").Value = "No students found."
    Else
        ss.Range("D3").Value = (outR - SEARCH_FIRST_ROW) & " record(s) found."
    End If
    If Not gSilent Then
        ss.Activate
        ss.Range("A" & SEARCH_FIRST_ROW).Select
    End If
    gLastMessage = CStr(ss.Range("D3").Value)
End Sub

Public Sub GoToDashboard()
    Worksheets(SH_DASH).Activate
    Range("inProgram").Select
End Sub

'------------------------------------------------------------------------------
'  INSTALL: creates the buttons (run once after importing this module)
'------------------------------------------------------------------------------
Public Sub InstallUpKidzSystem()
    Dim ds As Worksheet, ss As Worksheet
    Set ds = Worksheets(SH_DASH)
    Set ss = Worksheets(SH_SEARCH)

    DeleteButtons ds
    DeleteButtons ss
    ds.Range("B18").Value = ""

    ' New-enrollment buttons (row 18)
    AddButton ds, ds.Range("B18:C18"), "btnSubmit", "SUBMIT ENROLLMENT", "SubmitEnrollment", RGB(46, 125, 50), 13
    AddButton ds, ds.Range("D18"), "btnUpdate", "UPDATE ENROLLMENT", "UpdateEnrollment", RGB(30, 136, 229), 11
    AddButton ds, ds.Range("E18"), "btnClear", "CLEAR FORM", "ClearForm", RGB(120, 144, 156), 11
    ' Manage buttons (row 8)
    AddButton ds, ds.Range("G8:H8"), "btnLoad", "LOAD FOR EDIT", "LoadEnrollmentForEdit", RGB(30, 136, 229), 10
    AddButton ds, ds.Range("I8:J8"), "btnCancel", "CANCEL ENROLLMENT", "CancelEnrollment", RGB(198, 40, 40), 10
    ' Search button (row 12)
    AddButton ds, ds.Range("G12:J12"), "btnSearch", "SEARCH STUDENT", "SearchStudent", RGB(0, 137, 123), 10
    ' Back button on search sheet
    AddButton ss, ss.Range("H3:J3"), "btnBack", "BACK TO DASHBOARD", "GoToDashboard", RGB(55, 71, 79), 9

    ds.Activate
    Notify "UpKidz Enrollment System installed." & vbCrLf & vbCrLf & _
           "IMPORTANT: Save the workbook as 'Excel Macro-Enabled Workbook (*.xlsm)' " & _
           "so the macros and buttons are kept.", vbInformation
End Sub

'==============================================================================
'  FORM READING & VALIDATION
'==============================================================================
Private Function ReadForm() As Boolean
    Dim v As Variant, problems As String, pkgMonths As Long

    ReadForm = False
    fProgram = Trim$(CStr(Range("inProgram").Value))
    fName = CleanText(CStr(Range("inName").Value))
    fSchedule = CleanText(CStr(Range("inSchedule").Value))
    fPackage = Trim$(CStr(Range("inPackage").Value))
    fStartMonth = MonthIndex(CStr(Range("inStartMonth").Value))

    If SectionCol(fProgram) = 0 Then problems = problems & "- Select a valid Program." & vbCrLf
    If fName = "" Then problems = problems & "- Enter the student's Full Name." & vbCrLf
    If fSchedule = "" Then problems = problems & "- Enter the Schedule." & vbCrLf
    If fPackage = "" Then problems = problems & "- Select a Package." & vbCrLf
    If fStartMonth = 0 Then problems = problems & "- Select a Start Month (January to December)." & vbCrLf

    v = Range("inDuration").Value
    If IsNumeric(v) And Not IsEmpty(v) And CStr(v) <> "" Then
        If CDbl(v) = Int(CDbl(v)) And CDbl(v) >= 1 And CDbl(v) <= 12 Then
            fDuration = CLng(v)
        Else
            problems = problems & "- Duration must be a whole number from 1 to 12." & vbCrLf
        End If
    Else
        problems = problems & "- Enter the Duration (number of months, 1 to 12)." & vbCrLf
    End If

    v = Range("inFee").Value
    If IsNumeric(v) And Not IsEmpty(v) And CStr(v) <> "" Then
        fFee = CDbl(v)
        If fFee <= 0 Then problems = problems & "- Program Fee must be greater than zero." & vbCrLf
    Else
        problems = problems & "- Enter the Program Fee (numbers only)." & vbCrLf
    End If

    v = Range("inDiscount").Value
    If IsEmpty(v) Or CStr(v) = "" Then
        fDiscount = 0
    ElseIf IsNumeric(v) Then
        fDiscount = CDbl(v)
        If fDiscount < 0 Then problems = problems & "- Discount cannot be negative." & vbCrLf
    Else
        problems = problems & "- Discount must be a number (leave blank if none)." & vbCrLf
    End If

    If problems <> "" Then
        Notify "The enrollment cannot be submitted yet:" & vbCrLf & vbCrLf & problems, vbExclamation
        Exit Function
    End If

    If fDiscount > fFee Then
        Notify "The discount (" & Peso(fDiscount) & ") is greater than the program fee (" & _
               Peso(fFee) & ")." & vbCrLf & "Please correct the discount.", vbExclamation
        Exit Function
    End If
    fNet = fFee - fDiscount

    ' package must belong to the selected program
    pkgMonths = PackageMonths(fProgram, fPackage)
    If pkgMonths = -1 Then
        Notify "Package '" & fPackage & "' is not listed for " & fProgram & "." & vbCrLf & _
               "Choose a package from the dropdown (packages are maintained on the Settings sheet).", vbExclamation
        Exit Function
    End If
    If pkgMonths > 0 And pkgMonths <> fDuration Then
        If Ask("The package '" & fPackage & "' normally covers " & pkgMonths & " month(s), " & _
               "but Duration is " & fDuration & " month(s)." & vbCrLf & vbCrLf & _
               "Continue with " & fDuration & " month(s)?", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then
            Exit Function
        End If
    End If
    ReadForm = True
End Function

' Blocks accidental duplicates: same student already in this program section
' in any of the target months (ignoring the enrollment being edited).
Private Function CheckDuplicates(months As Variant, excludeID As String) As Boolean
    Dim i As Long, r As Long, ws As Worksheet, sc As Long, ic As Long, hits As String
    sc = SectionCol(fProgram)
    ic = IdCol(fProgram)
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
        CheckDuplicates = (Ask("POSSIBLE DUPLICATE" & vbCrLf & vbCrLf & fName & " is already enrolled in " & _
            SectionTitle(fProgram) & " for: " & hits & "." & vbCrLf & vbCrLf & _
            "Add this enrollment anyway? (Choose No unless this is a different student with the same name.)", _
            vbYesNo + vbExclamation + vbDefaultButton2) = vbYes)
    End If
End Function

' Ensures every target month has a free row BEFORE anything is written,
' so an enrollment is never half-recorded.
Private Function CheckCapacity(months As Variant, excludeID As String) As Boolean
    Dim i As Long, r As Long, ws As Worksheet, sc As Long, ic As Long, freeRows As Long, full As String
    sc = SectionCol(fProgram)
    ic = IdCol(fProgram)
    For i = LBound(months) To UBound(months)
        Set ws = Worksheets(MonthSheet(months(i)))
        freeRows = 0
        For r = FIRST_ROW To LAST_ROW
            If RowIsEmpty(ws, r, fProgram) Then
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
        Notify "No free rows left in " & SectionTitle(fProgram) & " for: " & Left$(full, Len(full) - 2) & "." & _
               vbCrLf & "Nothing was saved. Increase the number of rows or remove old records first.", vbCritical
        CheckCapacity = False
    End If
End Function

'==============================================================================
'  MONTHLY SHEET OPERATIONS
'==============================================================================
Private Sub WriteToMonths(months As Variant, id As String)
    Dim i As Long, ws As Worksheet, r As Long, sc As Long, price As Double
    sc = SectionCol(fProgram)
    If LCase$(CStr(Range("cfgPriceMode").Value)) = "split per month" Then
        price = Round(fNet / fDuration, 2)
    Else
        price = fNet
    End If
    For i = LBound(months) To UBound(months)
        Set ws = Worksheets(MonthSheet(months(i)))
        r = NextFreeRow(ws, fProgram)
        If r = 0 Then Err.Raise vbObjectError + 513, , "No free row in " & ws.Name   ' guarded by CheckCapacity
        ws.Cells(r, sc + 1).Value = fName
        ws.Cells(r, sc + 2).Value = fSchedule
        ws.Cells(r, sc + 3).Value = fPackage
        If SectionHasDiscount(fProgram) Then
            ws.Cells(r, sc + 4).Value = fDiscount
            ws.Cells(r, sc + 5).Value = price
        Else
            ws.Cells(r, sc + 4).Value = price
        End If
        ws.Cells(r, IdCol(fProgram)).Value = id
    Next i
End Sub

' Removes every row carrying this ID from all sections of all 12 months and
' moves the remaining rows up so each list stays continuous. Only values are
' moved (Name..ID); the numbering formulas and formatting stay untouched.
Private Function RemoveFromMonths(id As String) As Long
    Dim m As Long, k As Long, prg As String, ws As Worksheet
    Dim sc As Long, nCols As Long, data As Variant, outArr() As Variant, keep() As Variant
    Dim r As Long, c As Long, outR As Long, removed As Long, nRows As Long, found As Boolean

    nRows = LAST_ROW - FIRST_ROW + 1
    For m = 1 To 12
        Set ws = Worksheets(MonthSheet(m))
        For k = 1 To 3
            prg = ProgramByIndex(k)
            sc = SectionCol(prg)
            nCols = IdCol(prg) - sc            ' Name .. ID
            data = ws.Range(ws.Cells(FIRST_ROW, sc + 1), ws.Cells(LAST_ROW, IdCol(prg))).Value
            found = False
            For r = 1 To nRows
                If CStr(data(r, nCols)) = id Then found = True: Exit For
            Next r
            If found Then
                ReDim outArr(1 To nRows, 1 To nCols)
                outR = 0
                For r = 1 To nRows
                    If CStr(data(r, nCols)) = id Then
                        removed = removed + 1
                    ElseIf Not RowArrayEmpty(data, r, nCols) Then
                        outR = outR + 1
                        For c = 1 To nCols
                            outArr(outR, c) = data(r, c)
                        Next c
                    End If
                Next r
                ' clear the whole block (values only - formatting stays), then write back kept rows
                ws.Range(ws.Cells(FIRST_ROW, sc + 1), ws.Cells(LAST_ROW, IdCol(prg))).ClearContents
                If outR > 0 Then
                    ReDim keep(1 To outR, 1 To nCols)
                    For r = 1 To outR
                        For c = 1 To nCols
                            keep(r, c) = outArr(r, c)
                        Next c
                    Next r
                    ws.Range(ws.Cells(FIRST_ROW, sc + 1), ws.Cells(FIRST_ROW + outR - 1, IdCol(prg))).Value = keep
                End If
            End If
        Next k
    Next m
    RemoveFromMonths = removed
End Function

Private Function NextFreeRow(ws As Worksheet, prg As String) As Long
    Dim r As Long
    For r = FIRST_ROW To LAST_ROW
        If RowIsEmpty(ws, r, prg) Then
            NextFreeRow = r
            Exit Function
        End If
    Next r
    NextFreeRow = 0
End Function

' A row is free only if EVERY data cell (Name..ID) is empty - never overwrite.
Private Function RowIsEmpty(ws As Worksheet, r As Long, prg As String) As Boolean
    Dim c As Long
    For c = SectionCol(prg) + 1 To IdCol(prg)
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
'  ENROLLMENT HISTORY
'==============================================================================
Private Function NextEnrollmentID() As String
    Dim lg As Worksheet, r As Long, s As String, n As Long, mx As Long
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        s = CStr(lg.Cells(r, 1).Value)
        If s = "" Then Exit For
        If UCase$(Left$(s, 4)) = "UPK-" Then
            If IsNumeric(Mid$(s, 5)) Then
                n = CLng(Mid$(s, 5))
                If n > mx Then mx = n
            End If
        End If
    Next r
    NextEnrollmentID = "UPK-" & Format$(mx + 1, "0000")
End Function

Private Function NextLogRow() As Long
    Dim lg As Worksheet, r As Long
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        If CStr(lg.Cells(r, 1).Value) = "" Then
            NextLogRow = r
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 514, , "Enrollment History is full (" & (LOG_LAST_ROW - LOG_FIRST_ROW + 1) & " rows)."
End Function

Private Function FindLogRow(ByVal id As String) As Long
    Dim lg As Worksheet, r As Long, s As String
    Set lg = Worksheets(SH_LOG)
    For r = LOG_FIRST_ROW To LOG_LAST_ROW
        s = CStr(lg.Cells(r, 1).Value)
        If s = "" Then Exit For
        If StrComp(s, id, vbTextCompare) = 0 Then
            FindLogRow = r
            Exit Function
        End If
    Next r
    FindLogRow = 0
End Function

Private Function LogStatus(ByVal id As String) As String
    Dim r As Long
    r = FindLogRow(id)
    If r = 0 Then
        LogStatus = "Not in history"
    Else
        LogStatus = CStr(Worksheets(SH_LOG).Cells(r, 13).Value)
    End If
End Function

Private Sub LogNew(id As String, months As Variant)
    Dim lg As Worksheet, r As Long
    Set lg = Worksheets(SH_LOG)
    r = NextLogRow()
    lg.Cells(r, 1).Value = id
    lg.Cells(r, 2).Value = Now
    WriteLogDetails r, months
    lg.Cells(r, 13).Value = "Active"
    lg.Cells(r, 14).Value = Now
    lg.Cells(r, 15).Value = "Submitted."
End Sub

Private Sub LogUpdate(r As Long, months As Variant)
    Dim lg As Worksheet
    Set lg = Worksheets(SH_LOG)
    WriteLogDetails r, months
    lg.Cells(r, 13).Value = "Active"
    lg.Cells(r, 14).Value = Now
    lg.Cells(r, 15).Value = Trim$(lg.Cells(r, 15).Value & " Edited " & Format$(Now, "yyyy-mm-dd") & ".")
End Sub

Private Sub WriteLogDetails(r As Long, months As Variant)
    Dim lg As Worksheet
    Set lg = Worksheets(SH_LOG)
    lg.Cells(r, 3).Value = fProgram
    lg.Cells(r, 4).Value = fName
    lg.Cells(r, 5).Value = fSchedule
    lg.Cells(r, 6).Value = fPackage
    lg.Cells(r, 7).Value = MonthSheet(fStartMonth)
    lg.Cells(r, 8).Value = fDuration
    lg.Cells(r, 9).Value = MonthListText(months)
    lg.Cells(r, 10).Value = fFee
    lg.Cells(r, 11).Value = fDiscount
    lg.Cells(r, 12).Value = fNet
End Sub

'==============================================================================
'  HELPERS
'==============================================================================
' Months covered, wrapping December -> January (e.g. Nov + 3 = Nov, Dec, Jan)
Private Function CoveredMonths(startM As Long, dur As Long) As Variant
    Dim a() As Long, i As Long
    ReDim a(1 To dur)
    For i = 1 To dur
        a(i) = ((startM - 1 + i - 1) Mod 12) + 1
    Next i
    CoveredMonths = a
End Function

Private Function MonthSheet(ByVal m As Long) As String
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

Private Function ProgramByIndex(ByVal k As Long) As String
    ProgramByIndex = Choose(k, PRG_TODDLER, PRG_PRESCHOOL, PRG_TUTOR)
End Function

Private Function SectionCol(ByVal prg As String) As Long
    Select Case LCase$(prg)
        Case LCase$(PRG_TODDLER): SectionCol = COL_TODDLER
        Case LCase$(PRG_PRESCHOOL): SectionCol = COL_PRESCHOOL
        Case LCase$(PRG_TUTOR): SectionCol = COL_TUTOR
        Case Else: SectionCol = 0
    End Select
End Function

Private Function SectionHasDiscount(ByVal prg As String) As Boolean
    SectionHasDiscount = (LCase$(prg) <> LCase$(PRG_TUTOR))
End Function

' Hidden Enrollment-ID column = the column right after Price
Private Function IdCol(ByVal prg As String) As Long
    If SectionHasDiscount(prg) Then
        IdCol = SectionCol(prg) + 6
    Else
        IdCol = SectionCol(prg) + 5
    End If
End Function

Private Function SectionTitle(ByVal prg As String) As String
    Select Case LCase$(prg)
        Case LCase$(PRG_TODDLER): SectionTitle = "Toddlers Group"
        Case LCase$(PRG_PRESCHOOL): SectionTitle = "Preschool Group"
        Case LCase$(PRG_TUTOR): SectionTitle = "One-on-One Tutoring"
        Case Else: SectionTitle = prg
    End Select
End Function

' Months for a program/package from the Settings package table.
' Returns -1 if the package is not listed for that program, 0 if months blank.
Private Function PackageMonths(ByVal prg As String, ByVal pkg As String) As Long
    Dim progs As Variant, names As Variant, mos As Variant, i As Long
    progs = Range("PkgProgram").Value
    names = Range("PkgName").Value
    mos = Range("PkgMonths").Value
    For i = 1 To UBound(progs, 1)
        If StrComp(Trim$(CStr(progs(i, 1))), prg, vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(names(i, 1))), pkg, vbTextCompare) = 0 Then
            If IsNumeric(mos(i, 1)) And CStr(mos(i, 1)) <> "" Then
                PackageMonths = CLng(mos(i, 1))
            Else
                PackageMonths = 0
            End If
            Exit Function
        End If
    Next i
    PackageMonths = -1
End Function

Private Function FormSummary(months As Variant) As String
    FormSummary = "Program: " & fProgram & "  (" & SectionTitle(fProgram) & ")" & vbCrLf & _
                  "Full Name: " & fName & vbCrLf & _
                  "Schedule: " & fSchedule & vbCrLf & _
                  "Package: " & fPackage & vbCrLf & _
                  "Months (" & fDuration & "): " & MonthListText(months) & vbCrLf & _
                  "Program Fee: " & Peso(fFee) & vbCrLf & _
                  "Discount: " & Peso(fDiscount) & vbCrLf & _
                  "Net Price: " & Peso(fNet) & vbCrLf
End Function

Private Function Peso(ByVal x As Double) As String
    Peso = ChrW(8369) & Format$(x, "#,##0.00")
End Function

' Trim and collapse repeated spaces
Private Function CleanText(ByVal s As String) As String
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = s
End Function

Private Sub ClearInputs()
    Range("inProgram").Value = ""
    Range("inName").Value = ""
    Range("inSchedule").Value = ""
    Range("inPackage").Value = ""
    Range("inStartMonth").Value = ""
    Range("inDuration").Value = ""
    Range("inFee").Value = ""
    Range("inDiscount").Value = ""
End Sub

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

Private Sub AddButton(ws As Worksheet, target As Range, shpName As String, caption As String, _
                      macroName As String, fillColor As Long, fontSize As Single)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, target.Left + 3, target.Top + 2, target.Width - 6, target.Height - 4) ' 5 = rounded rectangle
    With shp
        .Name = shpName
        .OnAction = macroName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = 0
        .Placement = 2  ' move but don't size with cells
        With .TextFrame2
            .VerticalAnchor = 3                    ' middle
            .TextRange.Text = caption
            .TextRange.ParagraphFormat.Alignment = 2 ' center
            .TextRange.Font.Bold = -1
            .TextRange.Font.Size = fontSize
            .TextRange.Font.Name = "Arial"
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
End Sub
