Attribute VB_Name = "modFileUtils"
Option Explicit

' === modFileUtils � Fil- och mappfunktioner f�r indexerade filer ===

' Returnerar basnamn utan fil�ndelse
Public Function GetBaseName(filePath As String) As String
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
GetBaseName = fso.GetBaseName(filePath)
End Function

' Returnerar endast mappnamnet fr�n en full s�kv�g
Public Function GetFolderName(path As String) As String
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
GetFolderName = fso.GetParentFolderName(path)
End Function

' Skapar mapp om den inte finns (anv�nds t.ex. f�r alternativ PDF-s�kv�g)
Public Sub EnsureFolderExists(folderPath As String)
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(folderPath) Then
fso.CreateFolder folderPath
End If
End Sub

' Skapar en loggfil i angiven s�kv�g, returnerar fil-ID
Public Function CreateLogFile(logPath As String) As Integer
Dim fileNum As Integer
fileNum = FreeFile
Open logPath For Output As #fileNum
CreateLogFile = fileNum
End Function

' L�gger till rad i �ppen loggfil (med tidsst�mpel)
Public Sub AppendToLog(fileNum As Integer, message As String)
Print #fileNum, Format(Now, "yyyy-mm-dd HH\:nn\:ss") & " - " & message
End Sub


