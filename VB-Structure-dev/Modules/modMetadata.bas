Attribute VB_Name = "modMetaData"

' ========================================
' modMetadata – Metadatahantering för Word
' ========================================

Option Explicit

' === Läs in metadata från en fil ===
Public Function ReadMetadata(filePath As String) As FileMetadata
    Dim meta As FileMetadata
    Dim doc As Document

    Set doc = Documents.Open(filePath, ReadOnly:=True, Visible:=False)
    On Error Resume Next
    With doc.BuiltInDocumentProperties
        meta.Title = .Item("Title")
        meta.Subject = .Item("Subject")
        meta.Author = .Item("Author")
        meta.Keywords = .Item("Keywords")
        meta.DocumentDate = Format(.Item("Last Save Time"), "yyyy-mm-dd")
    End With
    On Error GoTo 0
    Call ExtractExtraMetadata(doc, meta)
    Call ExtractDocumentDateFromContent(doc, meta)
    doc.Close SaveChanges:=False

    ReadMetadata = meta
End Function

' === Skriv metadata till en fil ===
Public Sub WriteMetadata(filePath As String, meta As FileMetadata)
    Dim doc As Document
    Set doc = Documents.Open(filePath, ReadOnly:=False, Visible:=False)
    On Error Resume Next
    With doc.BuiltInDocumentProperties
        .Item("Title") = meta.Title
        .Item("Subject") = meta.Subject
        .Item("Author") = meta.Author
        .Item("Keywords") = meta.Keywords
    End With
    On Error GoTo 0
    doc.Close SaveChanges:=True
End Sub

' === Försök att hämta första YYYY-MM-DD i dokumentet ===
Private Sub ExtractDocumentDateFromContent(doc As Document, ByRef meta As FileMetadata)
    Dim re As Object
    Dim match As Object
    Dim text As String

    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "\b[0-9]{4}-[0-9]{2}-[0-9]{2}\b"
    re.Global = False

    text = doc.Content.text
    If re.Test(text) Then
        Set match = re.Execute(text)(0)
        meta.DocumentDate = match.Value
    End If
End Sub

' === Hämta metadata för alla markerade filer ===
Public Function CollectSelectedMetadata() As FileMetadata()
    Dim files() As IndexedFile
    files = GetSelectedIndexedFiles()

    Dim result() As FileMetadata
    Dim i As Long

    If (Not Not files) = False Then
        ReDim result(0)
        CollectSelectedMetadata = result
        Exit Function
    End If

    ReDim result(UBound(files))
    For i = 0 To UBound(files)
        result(i) = ReadMetadata(files(i).filePath)
    Next i

    CollectSelectedMetadata = result
End Function

' === Uppdatera globala MetadataIndex ===
Public Sub UpdateMetadataIndex()
    MetadataIndex = CollectSelectedMetadata()
End Sub

' === Skriv metadata till loggfil ===
Public Sub LogSelectedMetadata(logPath As String)
    Dim metas() As FileMetadata
    Dim files() As IndexedFile
    Dim i As Long
    Dim fNum As Integer

    metas = CollectSelectedMetadata()
    files = GetSelectedIndexedFiles()

    fNum = CreateLogFile(logPath)
    AppendToLog fNum, "File" & vbTab & "Title" & vbTab & "Subject" & vbTab & _
        "Author" & vbTab & "Keywords" & vbTab & "Date"

    For i = 0 To UBound(metas)
        AppendToLog fNum, files(i).fileName & vbTab & metas(i).Title & vbTab & _
            metas(i).Subject & vbTab & metas(i).Author & vbTab & metas(i).Keywords & _
            vbTab & metas(i).DocumentDate
    Next i

    Close #fNum
End Sub

' === Extrahera metadata från innehållet ===
Private Sub ExtractExtraMetadata(doc As Document, ByRef meta As FileMetadata)
    Dim labels As Variant
    Dim lbl As Variant
    Dim rng As Range
    Dim val As String

    labels = Array("Datum", "Handläggare", "Konstruktör")
    For Each lbl In labels
        Set rng = doc.Content
        With rng.Find
            .text = lbl & " *"
            .Forward = True
            .MatchWildcards = True
            If .Execute Then
                val = Trim(Replace(Split(rng.text, vbCr)(0), lbl, ""))
                Select Case lbl
                    Case "Datum": meta.Datum = val
                    Case "Handläggare": meta.Handlaggare = val
                    Case "Konstruktör": meta.Konstruktor = val
                End Select
            End If
        End With
    Next lbl
End Sub


