Attribute VB_Name = "modMain"
Option Explicit

Public cancelRequested As Boolean

Public Sub StartProcessing()
    Dim findTexts(1 To 5) As String
    Dim replaceTexts(1 To 5) As String
    Dim caseFlags(1 To 5) As Boolean
    Dim wordFlags(1 To 5) As Boolean
    Dim i As Long

    For i = 1 To 5
        findTexts(i) = frmReplaceTool.Controls("txtFind" & i).text
        replaceTexts(i) = frmReplaceTool.Controls("txtReplace" & i).text
        caseFlags(i) = frmReplaceTool.Controls("chkCase" & i).Value
        wordFlags(i) = frmReplaceTool.Controls("chkWhole" & i).Value
    Next i

    ProcessIndexedDocuments findTexts, replaceTexts, caseFlags, wordFlags, _
        frmReplaceTool.chkExportPDF.Value, frmReplaceTool.cmbPDFType.text, _
        frmReplaceTool.txtAltPDFPath.text, frmReplaceTool.txtPrefix.text, _
        frmReplaceTool.txtSuffix.text, frmReplaceTool.chkKeepOriginal.Value, _
        frmReplaceTool.cmbLanguage.text
End Sub



