Attribute VB_Name = "modBulleProof"
Sub AutoExec()
    ' D�ljer linjalen vid uppstart
    On Error Resume Next
    Application.ActiveWindow.DisplayRulers = False
    On Error GoTo 0
End Sub

Sub ViewRuler()
    ' Blockerar Ctrl+Shift+R och f�rs�k att visa linjalen
    MsgBox "�tkomst till linjalen �r permanent inaktiverad! Ring inte IT, anv�nd formatstilar...", vbExclamation
    On Error Resume Next
    Application.ActiveWindow.DisplayRulers = False
    On Error GoTo 0
End Sub





