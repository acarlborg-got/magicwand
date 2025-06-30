Attribute VB_Name = "modReplace"
Option Explicit

' === modReplace – Behandling av filer från index ===
Public Sub ProcessIndexedDocuments(findTexts() As String, replaceTexts() As String, _
    caseFlags() As Boolean, wordFlags() As Boolean, exportPDF As Boolean, exportPDFType As String, _
    altPDFPath As String, prefix As String, suffix As String, keepOriginal As Boolean, language As String)

    Dim fileList() As IndexedFile
    fileList = GetSelectedIndexedFiles()
    If (Not Not fileList) = False Then
        MsgBox "Inga filer har valts i index.", vbExclamation
        Exit Sub
    End If

    Dim i As Long, totalReplacements As Long, fileCount As Long
    Dim doc As Document
    Dim logFile As Integer, errorLogFile As Integer
    Dim logPath As String, errorLogPath As String
    Dim baseFolder As String

    baseFolder = GetBaseFolder(fileList)
    logPath = baseFolder & "\MagicWand_Log.txt"
    errorLogPath = baseFolder & "\MagicWand_Errors.txt"
    logFile = CreateLogFile(logPath)
    errorLogFile = CreateLogFile(errorLogPath)

    Call UpdateProgress(0)

    For i = 0 To UBound(fileList)
        DoEvents
        If cancelRequested Then Exit For

        On Error GoTo HandleError

        fileCount = fileCount + 1
        UpdateStatus "Processing", fileList(i).fileName, "Files: " & fileCount
        UpdateProgress fileCount / (UBound(fileList) + 1)

        Set doc = Documents.Open(fileList(i).filePath, ReadOnly:=False, Visible:=False)
                Call UpdateProgress((i + 1) / (UBound(fileList) + 1))
        Dim j As Long, replacementsInDoc As Long
        For j = 1 To 5
            If findTexts(j) <> "" Then
                Dim rep As Long
                rep = ReplaceAll(doc, findTexts(j), replaceTexts(j), caseFlags(j), wordFlags(j))
                totalReplacements = totalReplacements + rep
                replacementsInDoc = replacementsInDoc + rep
            End If
        Next j

        If keepOriginal Then
            Dim preservePath As String
            preservePath = fileList(i).filePath & "_originals"
            preservePath = Left(preservePath, InStrRev(preservePath, "\") - 1) & "\_originals"
            EnsureFolderExists preservePath
            doc.SaveAs2 preservePath & "\" & prefix & GetBaseName(fileList(i).fileName) & suffix & ".docx"
        Else
            doc.Save
        End If

        If exportPDF And (Not frmReplaceTool.chkExportPDFOnly.Value Or replacementsInDoc > 0) Then
            Dim exportPath As String, pdfPath As String
            exportPath = IIf(altPDFPath <> "", altPDFPath, Left(fileList(i).filePath, InStrRev(fileList(i).filePath, "\") - 1))
            EnsureFolderExists exportPath
            pdfPath = exportPath & "\" & prefix & GetBaseName(fileList(i).fileName) & suffix & ".pdf"

            doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
                CreateBookmarks:=wdExportCreateNoBookmarks, _
                UseISO19005_1:=IIf(exportPDFType = "PDF/A-1b", True, False)
        End If

        AppendToLog logFile, "Processed: " & fileList(i).filePath
        doc.Close SaveChanges:=False
ContinueLoop:
    Next i

Call UpdateProgress(1)

    Close #logFile
    Close #errorLogFile

    UpdateProgress 1
    UpdateStatus "Done", , "Processed: " & fileCount & " | Replacements: " & totalReplacements
    MsgBox "Behandling klar. Loggar sparade i: " & baseFolder, vbInformation
    Exit Sub

HandleError:
    If Not doc Is Nothing Then On Error Resume Next: doc.Close SaveChanges:=False
    AppendToLog errorLogFile, "Error in file: " & fileList(i).filePath & " - " & err.Description
    Resume ContinueLoop
End Sub

Private Function ReplaceAll(doc As Document, findText As String, replaceText As String, _
    caseSensitive As Boolean, wholeWord As Boolean) As Long

    Dim rng As Range, count As Long
    For Each rng In doc.StoryRanges
        Do
            count = count + ReplaceInRange(rng, findText, replaceText, caseSensitive, wholeWord)
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next

    Dim shp As Shape
    For Each shp In doc.Shapes
        If shp.TextFrame.HasText Then
            count = count + ReplaceInRange(shp.TextFrame.TextRange, findText, replaceText, caseSensitive, wholeWord)
        End If
    Next

    ReplaceAll = count
End Function

Private Function ReplaceInRange(rng As Range, findText As String, replaceText As String, _
    caseSensitive As Boolean, wholeWord As Boolean) As Long

    Dim count As Long, found As Boolean

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = caseSensitive
        .MatchWholeWord = wholeWord
    End With

    Do
        found = rng.Find.Execute
        If found Then
            Dim matchRange As Range
            Set matchRange = rng.Duplicate
            matchRange.text = replaceText
            count = count + 1
            rng.Start = matchRange.End
            rng.End = rng.StoryLength
        End If
    Loop While found

    ReplaceInRange = count
End Function

Private Function GetBaseFolder(files() As IndexedFile) As String
    If UBound(files) >= 0 Then
        GetBaseFolder = Left(files(0).filePath, InStrRev(files(0).filePath, "\") - 1)
    Else
        GetBaseFolder = ThisDocument.path
    End If
End Function


