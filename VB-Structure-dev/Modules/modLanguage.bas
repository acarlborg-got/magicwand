Attribute VB_Name = "modLanguage"
Option Explicit

Private LanguageDict As Object

' Loads translations from a .lng file located in the \languages folder
Public Sub LoadLanguage(languageCode As String)
    Dim f As Integer, line As String, pos As Integer
    Set LanguageDict = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler
    f = FreeFile
    Open ThisDocument.path & "\languages\" & languageCode & ".lng" For Input As #f
    Do Until EOF(f)
        Line Input #f, line
        pos = InStr(line, "=")
        If pos > 0 Then
            LanguageDict(Trim(Left(line, pos - 1))) = Trim(Mid(line, pos + 1))
        End If
    Loop
    Close #f
    Exit Sub
ErrHandler:
    ' If file missing, continue without translations
    On Error Resume Next
    Close #f
End Sub

' Returns the translation for a given key or the key itself if missing
Public Function T(key As String) As String
    If Not LanguageDict Is Nothing Then
        If LanguageDict.Exists(key) Then
            T = LanguageDict(key)
            Exit Function
        End If
    End If
    T = key
End Function
