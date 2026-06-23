Attribute VB_Name = "Module2_Projects"
Option Explicit

Public Sub Add_Project()
    Dim proj As String, phase As String
    proj = Trim(InputBox("New Project name:", "Add Project"))
    If Len(proj) = 0 Then Exit Sub
    phase = Trim(InputBox("Phase No. (optional):", "Add Project", "1"))
    Dim ls As Worksheet: Set ls = ThisWorkbook.Sheets("Lists")
    Dim r As Long: r = ls.Cells(ls.Rows.Count, 1).End(xlUp).Row + 1
    ls.Cells(r, 1).Value = proj
    ' seed a blank Current record
    Dim db As Worksheet: Set db = ThisWorkbook.Sheets("Database")
    Dim lastRow As Long: lastRow = db.Cells(db.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then lastRow = 4
    Dim nr As Long: nr = lastRow + 1
    Dim nid As Long
    If lastRow >= 5 Then nid = Application.WorksheetFunction.Max(db.Range("A5:A" & lastRow)) + 1 Else nid = 1
    db.Cells(nr, 1).Value = nid
    db.Cells(nr, 2).Value = proj
    db.Cells(nr, 3).Value = "Current"
    db.Cells(nr, 4).Value = Now
    db.Cells(nr, 5).Value = Environ("Username")
    db.Cells(nr, 6).Value = "seed (Add Project); Phase " & phase
    ThisWorkbook.Sheets("Cockpit").Range("SelectedProject").Value = proj
    MsgBox "Added project " & proj & " and selected it.", vbInformation
End Sub

Public Sub Validate_Inputs()
    Dim ck As Worksheet: Set ck = ThisWorkbook.Sheets("Cockpit")
    Dim msg As String
    If Len(Trim(ck.Range("SelectedProject").Value)) = 0 Then msg = msg & "- Project is blank" & vbCrLf
    If ck.Range("Units_Total").Value = 0 Then msg = msg & "- Unit mix has no units" & vbCrLf
    If Len(msg) = 0 Then MsgBox "Required inputs look complete.", vbInformation _
    Else MsgBox "Check these inputs:" & vbCrLf & vbCrLf & msg, vbExclamation
End Sub
