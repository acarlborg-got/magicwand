Attribute VB_Name = "modLauncher"


Option Explicit

Sub ReplaceTool()
    frmIndexBrowser.Show
End Sub

Sub ShowReplaceTool()
    ReplaceTool
End Sub

' Launch form by name if available
Public Sub ShowFormSafe(formName As String)
    On Error GoTo NotFound
    VBA.UserForms.Add(formName).Show
    Exit Sub
NotFound:
    MsgBox "Form '" & formName & "' not available.", vbInformation
End Sub


