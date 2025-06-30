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
        findTexts(i) = Me.Controls("txtFind" & i).text
        replaceTexts(i) = Me.Controls("txtReplace" & i).text
        caseFlags(i) = Me.Controls("chkCase" & i).Value
        wordFlags(i) = Me.Controls("chkWhole" & i).Value
    Next i

    ProcessIndexedDocuments findTexts, replaceTexts, caseFlags, wordFlags, _
        Me.chkExportPDF.Value, Me.cmbPDFType.text, Me.txtAltPDFPath.text, _
        Me.txtPrefix.text, Me.txtSuffix.text, Me.chkKeepOriginal.Value, _
        Me.cmbLanguage.text
End Sub



