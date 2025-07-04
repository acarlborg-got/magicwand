Attribute VB_Name = "modMetaTool"

' ========================================
' modMetaTool � Globala och regelstyrda metadata�tg�rder
' ========================================

Option Explicit

' === Applicera samma metadata p� alla valda filer ===
Public Sub ApplyGlobalMetadata(fieldName As String, valueSource As String, valueText As String, overwriteEmpty As Boolean)
    Dim files() As IndexedFile
    files = GetSelectedIndexedFiles()
    If (Not Not files) = False Then Exit Sub

    Dim i As Long
    Dim doc As Document
    Dim val As String
    For i = 0 To UBound(files)
        DoEvents
        Set doc = Documents.Open(files(i).filePath, ReadOnly:=False, Visible:=False)
        val = ResolveMetaValue(doc, valueSource, valueText)
        If Not overwriteEmpty Or GetDocProperty(doc, fieldName) = "" Then
            Call SetDocProperty(doc, fieldName, val)
        End If
        doc.Close SaveChanges:=True
    Next i
End Sub

' === Ber�kna v�rde baserat p� k�lla ===
Private Function ResolveMetaValue(doc As Document, source As String, text As String) As String
    Select Case source
        Case "Static Text": ResolveMetaValue = text
        Case "Current User": ResolveMetaValue = Application.UserName
        Case "Filename": ResolveMetaValue = GetBaseName(doc.Name)
        Case "Last Saved Date": ResolveMetaValue = Format(doc.BuiltInDocumentProperties("Last Save Time"), "yyyy-mm-dd")
        Case "Current Date": ResolveMetaValue = Format(Date, "yyyy-mm-dd")
        Case Else: ResolveMetaValue = text
    End Select
End Function

' === L�s v�rde fr�n dokumentets BuiltInProperty ===
Private Function GetDocProperty(doc As Document, propName As String) As String
    On Error Resume Next
    GetDocProperty = doc.BuiltInDocumentProperties(propName)
    On Error GoTo 0
End Function

' === S�tt BuiltInProperty ===
Private Sub SetDocProperty(doc As Document, propName As String, val As String)
    On Error Resume Next
    doc.BuiltInDocumentProperties(propName) = val
    On Error GoTo 0
End Sub

' === Applicera matris av regler (placeholder) ===
Public Sub ApplyMetadataMatrix(rules() As MetaRule)
    ' TODO: Implementera regelmotor
End Sub

