Attribute VB_Name = "modFindThings"
' === modFindThings – Datumanalys från indexerade filer ===
Option Explicit

Public Sub FindDatesInIndexedFiles()
    Dim fileList() As IndexedFile
    fileList = GetSelectedIndexedFiles()
    If (Not Not fileList) = False Then
        MsgBox "Inga filer har valts i index.", vbExclamation
        Exit Sub
    End If

    Dim dateDict As Object
    Set dateDict = CreateObject("Scripting.Dictionary")

    Dim i As Long, doc As Document
    Dim storyRng As Range, match As Object, matches As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = "\b(20\d{2})-(0[1-9]|1[0-2])-(0[1-9]|[12][0-9]|3[01])\b"

    ReDim MatchStats(0)
    frmReplaceTool.lstSpellingResult.Clear
    UpdateStatus "FindDates", , "Söker datum i dokument..."

    For i = 0 To UBound(fileList)
        DoEvents
        If cancelRequested Then Exit For

        Set doc = Documents.Open(fileList(i).filePath, ReadOnly:=True, Visible:=False)

        ' Statusuppdatering per fil
        UpdateStatus "Analyserar – " & fileList(i).fileName, , "Fil " & (i + 1) & " av " & (UBound(fileList) + 1)
        UpdateProgress (i + 1) / (UBound(fileList) + 1)

        For Each storyRng In doc.StoryRanges
            Do
                If regex.Test(storyRng.text) Then
                    Set matches = regex.Execute(storyRng.text)
                    For Each match In matches
                        If dateDict.Exists(match.Value) Then
                            dateDict(match.Value) = dateDict(match.Value) + 1
                        Else
                            dateDict.Add match.Value, 1
                        End If
                    Next match
                End If
                Set storyRng = storyRng.NextStoryRange
            Loop Until storyRng Is Nothing
        Next storyRng

        doc.Close SaveChanges:=False
    Next i

    UpdateProgress 1
    If dateDict.count = 0 Then
        UpdateStatus "FindDates", , "Inga datum hittades."
        MsgBox "Inga datum hittades i något dokument.", vbInformation
        Exit Sub
    End If

    ' Sortera resultat
    Dim keys() As String, values() As Long
    ReDim keys(0 To dateDict.count - 1)
    ReDim values(0 To dateDict.count - 1)

    For i = 0 To dateDict.count - 1
        keys(i) = dateDict.keys()(i)
        values(i) = dateDict.Items()(i)
    Next i

    Dim j As Long, tempK As String, tempV As Long
    For i = 0 To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(j) > values(i) Then
                tempV = values(i): values(i) = values(j): values(j) = tempV
                tempK = keys(i): keys(i) = keys(j): keys(j) = tempK
            End If
        Next j
    Next i

    ' Uppdatera MatchStats och formulär
    ReDim MatchStats(UBound(keys)) As tMatchStat
    frmReplaceTool.lstSpellingResult.Clear

    For i = 0 To UBound(keys)
        MatchStats(i).foundText = keys(i)
        MatchStats(i).count = values(i)
        MatchStats(i).sourceFile = "—"
        frmReplaceTool.lstSpellingResult.AddItem keys(i) & " (" & values(i) & ")"
    Next i

    ' Populera txtFind/Replace-fält
    For i = 0 To Min(4, UBound(keys))
        frmReplaceTool.Controls("txtFind" & (i + 1)).text = keys(i)
        frmReplaceTool.Controls("txtReplace" & (i + 1)).text = Format(Date, "yyyy-mm-dd")
    Next i

    UpdateStatus "FindDates complete", , "Unika datum: " & dateDict.count
    MsgBox "Datumanalys klar." & vbCrLf & "Unika datum: " & dateDict.count, vbInformation
End Sub

Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function
