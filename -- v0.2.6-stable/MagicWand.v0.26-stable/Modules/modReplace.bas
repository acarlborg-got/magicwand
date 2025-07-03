Attribute VB_Name = "modReplace"
Option Explicit

' #############################################
' ## modReplace – Updated with time tracking
' #############################################

Public Sub ProcessDocuments(folderPath As String, findTexts() As String, replaceTexts() As String, _
    caseFlags() As Boolean, wordFlags() As Boolean, exportPDF As Boolean, exportPDFType As String, _
    altPDFPath As String, prefix As String, suffix As String, includeSubfolders As Boolean, _
    keepOriginal As Boolean, language As String)

    Dim fso As Object, folder As Object
    Dim logFile As Integer, errorLogFile As Integer
    Dim logPath As String, errorLogPath As String
    Dim totalReplacements As Long, fileCount As Long
    Dim totalFiles As Long
    Dim localStartTime As Date: localStartTime = Now ' Lokal fallback för tid

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    logPath = folderPath & "\MagicWand_Log.txt"
    errorLogPath = folderPath & "\MagicWand_Errors.txt"
    logFile = CreateLogFile(logPath)
    errorLogFile = CreateLogFile(errorLogPath)

    totalFiles = CountWordFiles(folder, includeSubfolders)

    If totalFiles = 0 Then
        If includeSubfolders Then
            UpdateStatus "Info", , "No Word files found in root – continuing with subfolders..."
        Else
            UpdateStatus "No files", , "No Word files found in the selected folder."
            MsgBox "No Word files (.doc / .docx) found in the selected folder.", vbExclamation
            Close #logFile
            Close #errorLogFile
            Exit Sub
        End If
    End If

    ProcessFolder folder, findTexts, replaceTexts, caseFlags, wordFlags, exportPDF, exportPDFType, _
        altPDFPath, prefix, suffix, includeSubfolders, keepOriginal, logFile, errorLogFile, _
        totalReplacements, fileCount, totalFiles

    Close #logFile
    Close #errorLogFile

    UpdateProgress 1
    UpdateStatus "Done", , "Processed: " & fileCount & " files | Replacements: " & totalReplacements

    ' === Loggning till CSV ===
    Dim durationSeconds As Long
    If startTime = 0 Then
        durationSeconds = DateDiff("s", localStartTime, Now)
    Else
        durationSeconds = DateDiff("s", startTime, Now)
    End If

    Call LogAction("Replace+PDF", folderPath, includeSubfolders, exportPDF, exportPDFType, _
               altPDFPath, keepOriginal, fileCount, totalReplacements, fileCount, _
               durationSeconds, "Suffix=" & suffix)

    ' === Tidsbesparingsbedömning ===
    Dim estSavedSeconds As Double
    estSavedSeconds = EstimateTimeSaved("Replace+PDF", fileCount, totalReplacements, fileCount)

    Dim efficiencyMsg As String
    efficiencyMsg = "Processing complete. Logs saved in: " & folderPath & vbCrLf & vbCrLf & _
                    "? Estimated time saved: " & FormatTime(estSavedSeconds)

    MsgBox efficiencyMsg, vbInformation
End Sub



Public Sub ProcessFolder(folder As Object, findTexts() As String, replaceTexts() As String, _
    caseFlags() As Boolean, wordFlags() As Boolean, exportPDF As Boolean, _
    exportPDFType As String, altPDFPath As String, prefix As String, suffix As String, _
    includeSubfolders As Boolean, keepOriginal As Boolean, logFile As Integer, _
    errorLogFile As Integer, ByRef totalReplacements As Long, ByRef fileCount As Long, _
    ByVal totalFiles As Long)

    Dim file As Object, subFolder As Object
    Dim doc As Document, i As Integer
    Dim pdfPath As String, exportPath As String
    Dim replacementsInDoc As Long
    Dim preserveName As String
    preserveName = Trim(frmReplaceTool.txtPreserveSubFolder.Text)
    If preserveName = "" Then preserveName = "___NO_SUCH_FOLDER___" ' skydda vid tomt fält

    ' Loopa igenom filer (även om inga Wordfiler hittas i roten)
    For Each file In folder.files
        DoEvents
        If cancelRequested Then Exit Sub

        If LCase(Right(file.Name, 5)) = ".docx" Or LCase(Right(file.Name, 4)) = ".doc" Then
            On Error GoTo HandleError

            fileCount = fileCount + 1
            UpdateStatus "Processing", file.Name, "Files: " & fileCount & " | Replacements: " & totalReplacements
            UpdateProgress fileCount / totalFiles

            Set doc = Documents.Open(file.path, ReadOnly:=False, Visible:=False)
            replacementsInDoc = 0

            For i = 1 To 5
                If findTexts(i) <> "" Then
                    Dim rep As Long
                    rep = ReplaceAll(doc, findTexts(i), replaceTexts(i), caseFlags(i), wordFlags(i))
                    totalReplacements = totalReplacements + rep
                    replacementsInDoc = replacementsInDoc + rep
                End If
            Next i

            If keepOriginal Then
                Dim fullPreservePath As String
                fullPreservePath = folder.path & "\" & preserveName
                EnsureFolderExists fullPreservePath

                doc.SaveAs2 fileName:=fullPreservePath & "\" & prefix & GetBaseName(file.Name) & suffix & ".docx"
            Else
                doc.Save
            End If

            If exportPDF And (Not frmReplaceTool.chkExportPDFOnly.Value Or replacementsInDoc > 0) Then
                exportPath = IIf(altPDFPath <> "", altPDFPath, folder.path)
                EnsureFolderExists exportPath
                pdfPath = exportPath & "\" & prefix & GetBaseName(file.Name) & suffix & ".pdf"

                doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF, _
                    OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
                    CreateBookmarks:=wdExportCreateNoBookmarks, _
                    UseISO19005_1:=IIf(exportPDFType = "PDF/A-1b", True, False)
            End If

            AppendToLog logFile, "Processed: " & file.path
            doc.Close SaveChanges:=False
        End If
ContinueLoop:
    Next file

    ' Loopa submappar (även om inget behandlades i roten)
    If includeSubfolders Then
        For Each subFolder In folder.SubFolders
            If LCase(subFolder.Name) <> LCase(preserveName) Then
                ProcessFolder subFolder, findTexts, replaceTexts, caseFlags, wordFlags, exportPDF, _
                    exportPDFType, altPDFPath, prefix, suffix, includeSubfolders, keepOriginal, _
                    logFile, errorLogFile, totalReplacements, fileCount, totalFiles
            End If
        Next
    End If
    Exit Sub

HandleError:
    If Not doc Is Nothing Then On Error Resume Next: doc.Close SaveChanges:=False
    AppendToLog errorLogFile, "Error in file: " & file.path & " - " & err.Description
    Resume ContinueLoop
End Sub


Public Function ReplaceAll(doc As Document, findText As String, replaceText As String, _
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
        .Text = findText
        .Replacement.Text = ""
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
            matchRange.Text = replaceText
            count = count + 1
            rng.Start = matchRange.End
            rng.End = rng.StoryLength
        End If
    Loop While found

    ReplaceInRange = count
End Function

Private Function CountWordFiles(folder As Object, includeSubfolders As Boolean) As Long
    Dim f As Object, subFolder As Object, count As Long
    For Each f In folder.files
        If LCase(Right(f.Name, 5)) = ".docx" Or LCase(Right(f.Name, 4)) = ".doc" Then
            count = count + 1
        End If
    Next
    If includeSubfolders Then
        For Each subFolder In folder.SubFolders
            count = count + CountWordFiles(subFolder, True)
        Next
    End If
    CountWordFiles = count
End Function


