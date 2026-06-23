Attribute VB_Name = "Module2_Assets"
'================================================================
' Add_Asset  +  Validate_Inputs
'================================================================
Option Explicit

Public Sub Add_Asset()
    Dim pid As String, pname As String, site As String
    pid = Trim(InputBox("Asset / Phase ID (e.g. 3-2A):", "Add Asset"))
    If Len(pid) = 0 Then Exit Sub
    pname = Trim(InputBox("Phase Name:", "Add Asset", "Site - Bldg"))
    site = Trim(InputBox("Site (e.g. Site 3):", "Add Asset", "Site"))

    Dim wsL As Worksheet: Set wsL = ThisWorkbook.Sheets("Lists")
    Dim r As Long: r = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row + 1
    wsL.Cells(r, 1).Value = pid          ' PhaseID roster (drives both dropdowns)
    wsL.Cells(r, 2).Value = pname

    ' seed an initial blank/default record so look-ups resolve
    Dim wsDB As Worksheet: Set wsDB = ThisWorkbook.Sheets("Database")
    Dim lastRow As Long: lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then lastRow = 4
    Dim newRow As Long: newRow = lastRow + 1
    Dim newID As Long
    If lastRow >= 5 Then
        newID = Application.WorksheetFunction.Max(wsDB.Range("A5:A" & lastRow)) + 1
    Else
        newID = 1
    End If
    wsDB.Cells(newRow, 1).Value = newID
    wsDB.Cells(newRow, 2).Value = pid
    wsDB.Cells(newRow, 3).Value = pname
    wsDB.Cells(newRow, 4).Value = Now
    wsDB.Cells(newRow, 5).Value = Environ("Username")
    wsDB.Cells(newRow, 6).Value = "seed (Add Asset): " & site

    ThisWorkbook.Sheets("Cockpit").Range("SelectedPhase").Value = pid
    MsgBox "Added phase " & pid & " and selected it on the Cockpit.", vbInformation
End Sub

' Lightweight required-input check (conditional formatting also flags blanks).
Public Sub Validate_Inputs()
    Dim wsCk As Worksheet: Set wsCk = ThisWorkbook.Sheets("Cockpit")
    Dim msg As String
    If Len(Trim(wsCk.Range("SelectedPhase").Value)) = 0 Then msg = msg & "- Selected Phase is blank" & vbCrLf
    If wsCk.Range("Units_Total").Value = 0 Then msg = msg & "- Residential unit mix has no units" & vbCrLf
    If Len(msg) = 0 Then
        MsgBox "Required inputs look complete for the selected phase.", vbInformation
    Else
        MsgBox "Check these required inputs:" & vbCrLf & vbCrLf & msg, vbExclamation
    End If
End Sub
