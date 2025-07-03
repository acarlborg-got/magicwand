Attribute VB_Name = "modFileUtils"

Option Explicit

' Returnerar basnamn utan fil�ndelse
Function GetBaseName(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetBaseName = fso.GetBaseName(filePath)
End Function

' Returnerar mappens namn fr�n en full s�kv�g
Function GetFolderName(path As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
��� GetFolderName = fso.GetParentFolderName(path)
End Function

' Skapar mapp om den inte finns
Sub EnsureFolderExists(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

' Skapar en loggfil och returnerar filnummer
Public Function CreateLogFile(logPath As String) As Integer
    Dim fileNum As Integer
    fileNum = FreeFile
    Open logPath For Output As #fileNum
    CreateLogFile = fileNum
End Function

' L�gger till en rad i en loggfil
Sub AppendToLog(fileNum As Integer, message As String)
    Print #fileNum, Now & " - " & message
End Sub

