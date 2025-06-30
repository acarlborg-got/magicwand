Attribute VB_Name = "modUIUtils"


Option Explicit
' ##############################################
' ## Statusuppdatering i formulär
' ##############################################
Public Sub UpdateStatus(statusText As String, Optional progressText As String = "", Optional statsText As String = "")
    On Error Resume Next
    If VBA.UserForms.count > 0 Then
        With frmReplaceTool
            ' Om statusText är sammansatt i formatet: "Nivå – Filnamn | Meddelande"
            Dim state As String, fileName As String, message As String
            Dim parts() As String, subParts() As String
            
            If InStr(statusText, " – ") > 0 Or InStr(statusText, "|") > 0 Then
                ' Tolka uppdelat statusmeddelande
                parts = Split(statusText, "|")
                subParts = Split(parts(0), "–")
                state = Trim(subParts(0))
                If UBound(subParts) > 0 Then fileName = Trim(subParts(1))
                If UBound(parts) > 0 Then message = Trim(parts(1))
                
                .lblStatus2.Caption = state
                If fileName <> "" Then .lblProgress.Caption = fileName
                If message <> "" Then .lblStats.Caption = message
            Else
                ' Använd gamla formatet
                If statusText <> "" Then .lblStatus2.Caption = statusText
                If progressText <> "" Then .lblProgress.Caption = progressText
                If statsText <> "" Then .lblStats.Caption = statsText
            End If
            
            DoEvents
        End With
    End If
End Sub

' ##############################################
' ## Statusbar i formulär
' ##############################################
Public Sub UpdateProgress(progress As Double)
    On Error Resume Next
    If VBA.UserForms.count > 0 Then
        With frmReplaceTool
            If progress < 0 Then progress = 0
            If progress > 1 Then progress = 1
            .lblProgressBar.Width = 474 * progress ' Anpassa till faktisk bredd
            DoEvents
        End With
    End If
End Sub


