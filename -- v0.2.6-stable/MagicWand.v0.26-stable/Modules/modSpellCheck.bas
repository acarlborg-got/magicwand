Attribute VB_Name = "modSpellCheck"
' #############################################
' ## modSpellCheck – updated with time logging
' #############################################

Option Explicit

Public Sub PreScanSpelling(folderPath As String, language As String, _
                           includeSubfolders As Boolean, preserveFolderName As String)

    Dim t0 As Single: t0 = Timer

    Dim fallbackLangID As Long
    Select Case LCase(language)
        Case "svenska": fallbackLangID = 1053
        Case "english", "engelska": fallbackLangID = 2057
        Case Else: fallbackLangID = 2057
    End Select

    Dim fso As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    Dim spellingDict As Object
    Set spellingDict = CreateObject("Scripting.Dictionary")

    Dim fileCount As Long
    frmReplaceTool.lstSpellingResult.Clear
    UpdateStatus "Spellcheck", , "Scanning documents..."

    ScanSpellingRecursive folder, fallbackLangID, spellingDict, fileCount, _
                          includeSubfolders, preserveFolderName

    Dim count As Long: count = spellingDict.count

    If count = 0 Then
        UpdateStatus "Spellcheck complete", , "No misspellings found."
        UpdateProgress 1
        MsgBox "No spelling errors found in the scanned files.", vbInformation
        Exit Sub
    End If

    ' Sortera efter frekvens
    Dim keys() As String, values() As Long
    ReDim keys(0 To count - 1)
    ReDim values(0 To count - 1)

    Dim k As Long
    For k = 0 To count - 1
        keys(k) = spellingDict.keys()(k)
        values(k) = spellingDict.Items()(k)
    Next k

    Dim i As Long, j As Long, tempKey As String, tempVal As Long
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If values(j) > values(i) Then
                tempVal = values(i): values(i) = values(j): values(j) = tempVal
                tempKey = keys(i): keys(i) = keys(j): keys(j) = tempKey
            End If
        Next j
    Next i

    For i = 0 To Min(4, count - 1)
        frmReplaceTool.Controls("txtFind" & (i + 1)).Text = keys(i)
        frmReplaceTool.Controls("txtReplace" & (i + 1)).Text = GetSuggestion(keys(i), fallbackLangID)
        SetFieldOptionsForIndex keys(i), i + 1
    Next i

    Dim statsText As String
    statsText = "Files scanned: " & fileCount & " | Unique misspellings: " & count
    UpdateStatus "Spellcheck complete", , statsText

    ' Logg och effekt
    Dim seconds As Long: seconds = Timer - t0
    Call LogAction("Spellcheck", folderPath, includeSubfolders, False, "N/A", "", False, fileCount, 0, 0, seconds, "Lang=" & language)

    MsgBox "Spellcheck completed." & vbCrLf & statsText & vbCrLf & vbCrLf & _
           "? Estimated time saved: " & FormatTime(EstimateTimeSaved("Spellcheck", fileCount, 0, 0)), vbInformation

    Dim logPath As String
    logPath = folderPath & "\MagicWand_Spelling.txt"

    Dim fnum As Integer: fnum = FreeFile
    Open logPath For Output As #fnum
    Print #fnum, "MagicWand Spelling Log"
    Print #fnum, "Date: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    Print #fnum, "Scanned files: " & fileCount
    Print #fnum, "Unique misspellings: " & count
    Print #fnum, ""
    For i = 0 To count - 1
        Print #fnum, keys(i) & vbTab & values(i)
    Next i
    Close #fnum

    UpdateProgress 1
End Sub

Public Sub PreScanIsoDates(folderPath As String, includeSubfolders As Boolean, preserveFolderName As String)
    Dim t0 As Single: t0 = Timer

    Dim fso As Object, folder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    Dim spellingDict As Object
    Set spellingDict = CreateObject("Scripting.Dictionary")

    Dim fileCount As Long
    frmReplaceTool.lstSpellingResult.Clear
    UpdateStatus "DateScan", , "Söker datum..."

    ScanDatesRecursive folder, spellingDict, fileCount, includeSubfolders, preserveFolderName

    Dim count As Long: count = spellingDict.count
    If count = 0 Then
        UpdateStatus "DateScan complete", , "Inga datum hittades."
        UpdateProgress 1
        MsgBox "Inga datum hittades i valda filer.", vbInformation
        Exit Sub
    End If

    Dim keys() As String, values() As Long
    ReDim keys(0 To count - 1)
    ReDim values(0 To count - 1)

    Dim k As Long
    For k = 0 To count - 1
        keys(k) = spellingDict.keys()(k)
        values(k) = spellingDict.Items()(k)
    Next k

    Dim i As Long, j As Long, tempKey As String, tempVal As Long
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If values(j) > values(i) Then
                tempVal = values(i): values(i) = values(j): values(j) = tempVal
                tempKey = keys(i): keys(i) = keys(j): keys(j) = tempKey
            End If
        Next j
    Next i

    Dim todaysDate As String: todaysDate = Format(Date, "yyyy-mm-dd")
    For i = 0 To Min(4, count - 1)
        frmReplaceTool.Controls("txtFind" & (i + 1)).Text = keys(i)
        frmReplaceTool.Controls("txtReplace" & (i + 1)).Text = todaysDate
        SetFieldOptionsForIndex keys(i), i + 1
    Next i

    UpdateSpellingListLive spellingDict

    Dim statsText As String
    statsText = "Files scanned: " & fileCount & " | Unique dates: " & count
    UpdateStatus "DateScan complete", , statsText

    Dim seconds As Long: seconds = Timer - t0
    Call LogAction("FindDates", folderPath, includeSubfolders, False, "N/A", "", False, fileCount, 0, 0, seconds, "Lang=N/A")

    MsgBox "Datumscanning klar!" & vbCrLf & statsText & vbCrLf & vbCrLf & _
           "? Estimated time saved: " & FormatTime(EstimateTimeSaved("FindDates", fileCount, 0, 0)), vbInformation

    UpdateProgress 1
End Sub



Private Sub ScanDatesRecursive(folder As Object, ByRef dict As Object, ByRef fileCount As Long, _
                               includeSubfolders As Boolean, preserveFolderName As String)

    If LCase(folder.Name) = LCase(preserveFolderName) Then Exit Sub

    Dim file As Object, subFolder As Object
    Dim doc As Document
    Dim regEx As Object, matches, match, rng As Range
    Dim key As String

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.pattern = "\b(?:\d{4}|[Xx]{1,4})-(?:\d{2}|[Xx]{1,2})-(?:\d{2}|[Xx]{1,2})\b"
    regEx.Global = True
    regEx.IgnoreCase = True

    For Each file In folder.files
        If LCase(Right(file.Name, 5)) = ".docx" Or LCase(Right(file.Name, 4)) = ".doc" Then
            On Error GoTo SkipFile
            fileCount = fileCount + 1
            UpdateStatus "Söker datum: " & file.Name, , "Fil " & fileCount
            UpdateProgress 0.01 * (fileCount Mod 100)

            Set doc = Documents.Open(file.path, ReadOnly:=True, Visible:=False)

            ' === Sök i alla story-ranges (inkl. huvudtext, sidhuvud, sidfot, fotnoter etc) ===
            For Each rng In doc.StoryRanges
                Do
                    Set matches = regEx.Execute(rng.Text)
                    For Each match In matches
                        key = match.Value
                        If dict.Exists(key) Then
                            dict(key) = dict(key) + 1
                        Else
                            dict.Add key, 1
                        End If
                    Next match
                    Set rng = rng.NextStoryRange
                Loop Until rng Is Nothing
            Next rng

            doc.Close SaveChanges:=False
        End If
SkipFile:
        On Error GoTo 0
    Next

    If includeSubfolders Then
        For Each subFolder In folder.SubFolders
            ScanDatesRecursive subFolder, dict, fileCount, includeSubfolders, preserveFolderName
        Next
    End If
End Sub




Private Sub ScanSpellingRecursive(folder As Object, fallbackLangID As Long, _
                                  ByRef dict As Object, ByRef fileCount As Long, _
                                  includeSubfolders As Boolean, preserveFolderName As String)

    If LCase(folder.Name) = LCase(preserveFolderName) Then Exit Sub

    Dim file As Object, subFolder As Object
    Dim doc As Document, err As Object, wordKey As String

    For Each file In folder.files
        If LCase(Right(file.Name, 5)) = ".docx" Or LCase(Right(file.Name, 4)) = ".doc" Then
            On Error GoTo SkipFile
            fileCount = fileCount + 1
            UpdateStatus "Spellcheck – " & file.Name, , "File " & fileCount
            UpdateProgress 0.01 * (fileCount Mod 100)

            Set doc = Documents.Open(file.path, ReadOnly:=True, Visible:=False)
            doc.Content.LanguageID = fallbackLangID

            For Each err In doc.SpellingErrors
                wordKey = LCase(Trim(err.Text))
                If Len(wordKey) > 1 Then
                    If dict.Exists(wordKey) Then
                        dict(wordKey) = dict(wordKey) + 1
                    Else
                        dict.Add wordKey, 1
                    End If
                    UpdateSpellingListLive dict
                End If
            Next

            doc.Close SaveChanges:=False
        End If
SkipFile:
        On Error GoTo 0
    Next

    If includeSubfolders Then
        For Each subFolder In folder.SubFolders
            ScanSpellingRecursive subFolder, fallbackLangID, dict, fileCount, _
                                  includeSubfolders, preserveFolderName
        Next
    End If
End Sub

Private Sub UpdateSpellingListLive(dict As Object)
    Dim count As Long: count = dict.count
    If count = 0 Then Exit Sub

    Dim keys() As String, values() As Long
    ReDim keys(0 To count - 1)
    ReDim values(0 To count - 1)

    Dim i As Long, j As Long
    For i = 0 To count - 1
        keys(i) = dict.keys()(i)
        values(i) = dict.Items()(i)
    Next i

    Dim tempK As String, tempV As Long
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If values(j) > values(i) Then
                tempV = values(i): values(i) = values(j): values(j) = tempV
                tempK = keys(i): keys(i) = keys(j): keys(j) = tempK
            End If
        Next j
    Next i

    With frmReplaceTool.lstSpellingResult
        .Clear
        For i = 0 To Min(99, count - 1)
            .AddItem keys(i) & " (" & values(i) & ")"
        Next i
    End With
End Sub

Private Sub SetFieldOptionsForIndex(word As String, index As Long)
    Dim allText As String: allText = ActiveDocument.Content.Text

    If InStr(1, allText, LCase(word)) > 0 And InStr(1, allText, UCase(word)) > 0 Then
        frmReplaceTool.Controls("chkCase" & index).Value = False
    Else
        frmReplaceTool.Controls("chkCase" & index).Value = True
    End If

    If RegexMatch(allText, "[a-zA-ZåäöÅÄÖ]" & word & "[a-zA-ZåäöÅÄÖ]") Then
        frmReplaceTool.Controls("chkWhole" & index).Value = False
    Else
        frmReplaceTool.Controls("chkWhole" & index).Value = True
    End If
End Sub

Private Function GetSuggestion(word As String, langID As Long) As String
    Dim doc As Document
    Dim rng As Range
    Dim suggestion As String

    Set doc = Application.Documents.Add(Visible:=False)
    Set rng = doc.Range
    rng.Text = word

    rng.LanguageID = langID
    rng.Paragraphs(1).Range.LanguageID = langID

    If Not Application.CheckSpelling(rng.Text, , , langID) Then
        If doc.SpellingErrors.count > 0 Then
            Dim suggestions As Object
            Set suggestions = doc.SpellingErrors(1).GetSpellingSuggestions
            If Not suggestions Is Nothing Then
                If suggestions.count > 0 Then
                    suggestion = suggestions(1)
                End If
            End If
        End If
    End If

    doc.Close SaveChanges:=False

    If suggestion = "" Then
        GetSuggestion = word
    Else
        GetSuggestion = suggestion
    End If
End Function

Private Function RegexMatch(ByVal inputText As String, ByVal pattern As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    RegexMatch = re.Test(inputText)
End Function

Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function


