Attribute VB_Name = "Module1_Publish"
'================================================================
' Publish_Project  -  writes the blue Console inputs into the
' relational Database as a new "Current" record (prior Current ->
' Historical), then enforces the FIFO version cap.
' Database: row 4 headers, data from row 5.
'   A=RecordID B=Project C=Status D=Timestamp E=UpdatedBy F=Notes
'   first input field = column 7 (G).  68 fields.
'================================================================
Option Explicit
Private Const HDR As Long = 4
Private Const FIRST As Long = 5
Private Const FCOL As Long = 7
Private Const PROJNAME_ADDR As String = "$D$15"

Public Function FieldAddrs() As Variant
    Dim A(68) As String
    A(1) = "$D$14"
    A(2) = "$D$15"
    A(3) = "$D$16"
    A(4) = "$D$17"
    A(5) = "$D$18"
    A(6) = "$D$19"
    A(7) = "$D$21"
    A(8) = "$D$22"
    A(9) = "$D$23"
    A(10) = "$D$24"
    A(11) = "$D$26"
    A(12) = "$D$27"
    A(13) = "$D$28"
    A(14) = "$D$29"
    A(15) = "$D$30"
    A(16) = "$D$31"
    A(17) = "$D$32"
    A(18) = "$B$36"
    A(19) = "$C$36"
    A(20) = "$D$36"
    A(21) = "$E$36"
    A(22) = "$B$37"
    A(23) = "$C$37"
    A(24) = "$D$37"
    A(25) = "$E$37"
    A(26) = "$B$38"
    A(27) = "$C$38"
    A(28) = "$D$38"
    A(29) = "$E$38"
    A(30) = "$B$39"
    A(31) = "$C$39"
    A(32) = "$D$39"
    A(33) = "$E$39"
    A(34) = "$D$43"
    A(35) = "$D$44"
    A(36) = "$D$45"
    A(37) = "$D$46"
    A(38) = "$B$50"
    A(39) = "$D$50"
    A(40) = "$B$51"
    A(41) = "$D$51"
    A(42) = "$B$52"
    A(43) = "$D$52"
    A(44) = "$B$53"
    A(45) = "$D$53"
    A(46) = "$B$54"
    A(47) = "$D$54"
    A(48) = "$D$61"
    A(49) = "$D$62"
    A(50) = "$D$63"
    A(51) = "$D$64"
    A(52) = "$D$65"
    A(53) = "$D$66"
    A(54) = "$D$67"
    A(55) = "$D$68"
    A(56) = "$D$69"
    A(57) = "$D$78"
    A(58) = "$D$79"
    A(59) = "$D$80"
    A(60) = "$D$81"
    A(61) = "$D$82"
    A(62) = "$D$83"
    A(63) = "$D$84"
    A(64) = "$D$85"
    A(65) = "$D$86"
    A(66) = "$D$87"
    A(67) = "$D$88"
    A(68) = "$D$89"
    Dim out() As String, i As Long
    ReDim out(1 To 68)
    For i = 1 To 68: out(i) = A(i): Next i
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
    db.Rows(srcRow).Copy
    ar.Rows(dest).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub
