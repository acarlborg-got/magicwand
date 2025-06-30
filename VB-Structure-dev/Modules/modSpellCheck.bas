Attribute VB_Name = "modSpellCheck"
Option Explicit

Public Sub PreScanSpelling(filePaths() As String, language As String, preserveFolderName As String)
    Dim fallbackLangID As Long
    Select Case LCase(language)
        Case "svenska": fallbackLangID = 1053
        Case "english", "engelska": fallbackLangID = 2057
        Case Else: fallbackLangID = 2057
    End Select

    Dim spellingDict As Object
    Set spellingDict = CreateObject("Scripting.Dictionary")

    Dim fileCount As Long
    frmReplaceTool.lstSpellingResult.Clear
    UpdateStatus "Spellcheck", , "Scanning documents..."

    Dim i As Long
    For i = 0 To UBound(filePaths)
        If LCase(Right(filePaths(i), 5)) = ".docx" Or LCase(Right(filePaths(i), 4)) = ".doc" Then
            On Error GoTo SkipFile
            fileCount = fileCount + 1
            UpdateStatus "Spellcheck – " & filePaths(i), , "File " & fileCount
            UpdateProgress fileCount / (UBound(filePaths) + 1)

            Dim doc As Document
            Set doc = Documents.Open(filePaths(i), ReadOnly:=True, Visible:=False)
            doc.Content.LanguageID = fallbackLangID

            Dim spellingErrors As Object
            Set spellingErrors = doc.spellingErrors

            Dim err As Object, wordKey As String
            For Each err In spellingErrors
                wordKey = LCase(Trim(err.text))
                If Len(wordKey) > 1 Then
                    If spellingDict.Exists(wordKey) Then
                        spellingDict(wordKey) = spellingDict(wordKey) + 1
                    Else
                        spellingDict.Add wordKey, 1
                    End If
                End If
            Next

            doc.Close SaveChanges:=False
SkipFile:
            On Error GoTo 0
        End If
    Next i

    Dim count As Long: count = spellingDict.count
    If count = 0 Then
        UpdateStatus "Spellcheck complete", , "No misspellings found."
        UpdateProgress 1
        MsgBox "No spelling errors found in the scanned files.", vbInformation
        Exit Sub
    End If

    ' Sortera stavfel efter frekvens
    Dim keys() As String, values() As Long
    ReDim keys(0 To count - 1)
    ReDim values(0 To count - 1)

    Dim k As Long
    For k = 0 To count - 1
        keys(k) = spellingDict.keys()(k)
        values(k) = spellingDict.Items()(k)
    Next k

    Dim j As Long, tempKey As String, tempVal As Long
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If values(j) > values(i) Then
                tempVal = values(i): values(i) = values(j): values(j) = tempVal
                tempKey = keys(i): keys(i) = keys(j): keys(j) = tempKey
            End If
        Next j
    Next i

    ' Uppdatera lstSpellingResult
    With frmReplaceTool.lstSpellingResult
        .Clear
        For i = 0 To Min(99, count - 1)
            .AddItem keys(i) & " (" & values(i) & ")"
        Next i
    End With

    ' Fyll txtFindN, txtReplaceN, chkCaseN, chkWholeN
    For i = 0 To Min(4, count - 1)
        Dim word As String: word = keys(i)
        frmReplaceTool.Controls("txtFind" & (i + 1)).text = word
        frmReplaceTool.Controls("txtReplace" & (i + 1)).text = GetSuggestion(word, fallbackLangID)
        SetFieldOptionsForIndex word, i + 1
    Next i

    ' Klar
    Dim statsText As String
    statsText = "Files scanned: " & fileCount & " | Unique misspellings: " & count
    UpdateStatus "Spellcheck complete", , statsText
    MsgBox "Spellcheck completed." & vbCrLf & statsText, vbInformation

    Dim logPath As String
    logPath = Environ("TEMP") & "\MagicWand_Spelling.txt"
    Dim fNum As Integer: fNum = FreeFile
    Open logPath For Output As #fNum
    Print #fNum, "MagicWand Spelling Log"
    Print #fNum, "Date: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    Print #fNum, "Scanned files: " & fileCount
    Print #fNum, "Unique misspellings: " & count
    Print #fNum, ""
    For i = 0 To count - 1
        Print #fNum, keys(i) & vbTab & values(i)
    Next i
    Close #fNum

    UpdateProgress 1
End Sub

Private Sub SetFieldOptionsForIndex(word As String, index As Long)
    On Error Resume Next
    If InStr(word, " ") > 0 Then
        frmReplaceTool.Controls("chkWhole" & index).Value = False
    Else
        frmReplaceTool.Controls("chkWhole" & index).Value = True
    End If
    If word = UCase(word) Then
        frmReplaceTool.Controls("chkCase" & index).Value = True
    Else
        frmReplaceTool.Controls("chkCase" & index).Value = False
    End If
End Sub

Private Function GetSuggestion(word As String, langID As Long) As String
    Dim doc As Document
    Dim rng As Range
    Dim suggestion As String

    Set doc = Application.Documents.Add(Visible:=False)
    Set rng = doc.Range
    rng.text = word

    rng.LanguageID = langID
    rng.Paragraphs(1).Range.LanguageID = langID

    If Not Application.CheckSpelling(rng.text, , , langID) Then
        If doc.spellingErrors.count > 0 Then
            Dim suggestions As Object
            Set suggestions = doc.spellingErrors(1).GetSpellingSuggestions
            If Not suggestions Is Nothing And suggestions.count > 0 Then
                suggestion = suggestions(1)
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

Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function
