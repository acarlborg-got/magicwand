Attribute VB_Name = "modBulleProof"
Sub AutoExec()
    ' Döljer linjalen vid uppstart
    On Error Resume Next
    Application.ActiveWindow.DisplayRulers = False
    On Error GoTo 0
End Sub

Sub ViewRuler()
    ' Blockerar Ctrl+Shift+R och försök att visa linjalen
    MsgBox "Åtkomst till linjalen är permanent inaktiverad! Ring inte IT, använd formatstilar...", vbExclamation
    On Error Resume Next
    Application.ActiveWindow.DisplayRulers = False
    On Error GoTo 0
End Sub





