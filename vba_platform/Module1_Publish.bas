Attribute VB_Name = "Module1_Publish"
'================================================================
' Publish_Project  -  writes the blue Console inputs into the
' relational Database as a new "Current" record (prior Current ->
' Historical), then enforces the FIFO version cap.
' Database: row 4 headers, data from row 5.
'   A=RecordID B=Project C=Status D=Timestamp E=UpdatedBy F=Notes
'   first input field = column 7 (G).  69 fields.
'================================================================
Option Explicit
Private Const HDR As Long = 4
Private Const FIRST As Long = 5
Private Const FCOL As Long = 7
Private Const PROJNAME_ADDR As String = "$D$18"

Public Function FieldAddrs() As Variant
    Dim A(69) As String
    A(1) = "$D$17"
    A(2) = "$D$18"
    A(3) = "$D$19"
    A(4) = "$D$20"
    A(5) = "$D$21"
    A(6) = "$D$22"
    A(7) = "$D$23"
    A(8) = "$D$25"
    A(9) = "$D$26"
    A(10) = "$D$27"
    A(11) = "$D$28"
    A(12) = "$D$30"
    A(13) = "$D$31"
    A(14) = "$D$32"
    A(15) = "$D$33"
    A(16) = "$D$34"
    A(17) = "$D$35"
    A(18) = "$D$36"
    A(19) = "$B$40"
    A(20) = "$C$40"
    A(21) = "$D$40"
    A(22) = "$E$40"
    A(23) = "$B$41"
    A(24) = "$C$41"
    A(25) = "$D$41"
    A(26) = "$E$41"
    A(27) = "$B$42"
    A(28) = "$C$42"
    A(29) = "$D$42"
    A(30) = "$E$42"
    A(31) = "$B$43"
    A(32) = "$C$43"
    A(33) = "$D$43"
    A(34) = "$E$43"
    A(35) = "$D$47"
    A(36) = "$D$48"
    A(37) = "$D$49"
    A(38) = "$D$50"
    A(39) = "$B$54"
    A(40) = "$D$54"
    A(41) = "$B$55"
    A(42) = "$D$55"
    A(43) = "$B$56"
    A(44) = "$D$56"
    A(45) = "$B$57"
    A(46) = "$D$57"
    A(47) = "$B$58"
    A(48) = "$D$58"
    A(49) = "$D$65"
    A(50) = "$D$66"
    A(51) = "$D$67"
    A(52) = "$D$68"
    A(53) = "$D$69"
    A(54) = "$D$70"
    A(55) = "$D$71"
    A(56) = "$D$72"
    A(57) = "$D$73"
    A(58) = "$D$82"
    A(59) = "$D$83"
    A(60) = "$D$84"
    A(61) = "$D$85"
    A(62) = "$D$86"
    A(63) = "$D$87"
    A(64) = "$D$88"
    A(65) = "$D$89"
    A(66) = "$D$90"
    A(67) = "$D$91"
    A(68) = "$D$92"
    A(69) = "$D$93"
    Dim out() As String, i As Long
    ReDim out(1 To 69)
    For i = 1 To 69: out(i) = A(i): Next i
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
