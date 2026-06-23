Attribute VB_Name = "Module3_Support"
'================================================================
' Request_Support (ESCALATION) - books time with the model owner.
' Set CALENDLY_URL to your real scheduling link.
'================================================================
Option Explicit

Public Const CALENDLY_URL As String = "https://calendly.com/your-org/cockpit-model-support"

Public Sub Request_Support()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink CALENDLY_URL
    If Err.Number <> 0 Then
        MsgBox "Could not open the scheduling link. Please browse to:" & vbCrLf & _
               CALENDLY_URL, vbInformation
    End If
End Sub
