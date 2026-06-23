Attribute VB_Name = "Module3_Support"
Option Explicit
Public Const CALENDLY_URL As String = "https://calendly.com/your-org/model-support"
Public Sub Request_Support()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink CALENDLY_URL
    If Err.Number <> 0 Then MsgBox "Open: " & CALENDLY_URL, vbInformation
End Sub
