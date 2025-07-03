Attribute VB_Name = "modMain"
Option Explicit

Public startTime As Date
Public cancelRequested As Boolean

Public Sub StartProcessing()
    Dim folderPath As String, altPDFPath As String
    Dim exportPDF As Boolean, exportPDFType As String
    Dim exportPDFOnly As Boolean
    Dim prefix As String, suffix As String
    Dim language As String
    Dim includeSubfolders As Boolean, keepOriginal As Boolean

    Dim findTexts(1 To 5) As String
    Dim replaceTexts(1 To 5) As String
    Dim caseSensitive(1 To 5) As Boolean
    Dim wholeWord(1 To 5) As Boolean
    
    startTime = Now ' === För loggning

    With frmReplaceTool
        folderPath = .txtFolderPath.Text
        altPDFPath = .txtAltPDFPath.Text
        exportPDF = .chkExportPDF.Value
        exportPDFOnly = .chkExportPDFOnly.Value
        exportPDFType = .cmbPDFType.Text
        prefix = .txtPrefix.Text
        suffix = .txtSuffix.Text
        language = .cmbLanguage.Text
        includeSubfolders = .chkIncludeSubfolders.Value
        keepOriginal = .chkKeepOriginal.Value

        findTexts(1) = .txtFind1.Text
        replaceTexts(1) = .txtReplace1.Text
        caseSensitive(1) = .chkCase1.Value
        wholeWord(1) = .chkWhole1.Value

        findTexts(2) = .txtFind2.Text
        replaceTexts(2) = .txtReplace2.Text
        caseSensitive(2) = .chkCase2.Value
        wholeWord(2) = .chkWhole2.Value

        findTexts(3) = .txtFind3.Text
        replaceTexts(3) = .txtReplace3.Text
        caseSensitive(3) = .chkCase3.Value
        wholeWord(3) = .chkWhole3.Value

        findTexts(4) = .txtFind4.Text
        replaceTexts(4) = .txtReplace4.Text
        caseSensitive(4) = .chkCase4.Value
        wholeWord(4) = .chkWhole4.Value

        findTexts(5) = .txtFind5.Text
        replaceTexts(5) = .txtReplace5.Text
        caseSensitive(5) = .chkCase5.Value
        wholeWord(5) = .chkWhole5.Value
    End With

    cancelRequested = False

    ' Här skickar vi vidare exportPDFOnly via frmReplaceTool direkt – redan hanterat i modReplace
    ProcessDocuments folderPath, findTexts, replaceTexts, caseSensitive, wholeWord, _
        exportPDF, exportPDFType, altPDFPath, prefix, suffix, includeSubfolders, keepOriginal, language
End Sub

