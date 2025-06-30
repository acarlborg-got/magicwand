Attribute VB_Name = "modShared"
Option Explicit

Public Function GetIndexedFiles() As Variant
    GetIndexedFiles = selectedFiles  ' tFileSelection() fr�n frmIndexBrowser
End Function

' Returnerar urvalet av filer
Public Function GetSelectedFiles() As tFileSelection()
    GetSelectedFiles = selectedFiles
End Function

' Returnerar urvalet av mappar
Public Function GetSelectedFolders() As IndexedFolder()
    GetSelectedFolders = selectedFolders
End Function

' Returnerar fullst�ndiga s�kv�gar till valda filer
Public Function GetSelectedFilePaths() As String()
    Dim result() As String
    Dim i As Long, j As Long
    Dim fname As String, pid As Long
    Dim found As Boolean
    Dim fCount As Long

    If (Not Not selectedFiles) = False Then
        ReDim result(0)
        result(0) = ""
        GetSelectedFilePaths = result
        Exit Function
    End If

    fCount = UBound(selectedFiles) + 1
    ReDim result(fCount - 1)

    For i = 0 To UBound(selectedFiles)
        fname = selectedFiles(i).fileName
        pid = selectedFiles(i).parentID
        found = False

        For j = 0 To UBound(IndexedFiles)
            If IndexedFiles(j).fileName = fname And IndexedFiles(j).parentFolderID = pid Then
                result(i) = IndexedFiles(j).filePath
                found = True
                Exit For
            End If
        Next j

        If Not found Then
            result(i) = "MISSING: " & fname
        End If
    Next i

    GetSelectedFilePaths = result
End Function
' Returnerar hela IndexedFile-objekt f�r valda filer
Public Function GetSelectedIndexedFiles() As IndexedFile()
    Dim result() As IndexedFile
    Dim i As Long, j As Long
    Dim count As Long: count = 0

    If (Not Not selectedFiles) = False Then
        ReDim result(0)
        GetSelectedIndexedFiles = result
        Exit Function
    End If

    For i = 0 To UBound(selectedFiles)
        For j = 0 To UBound(IndexedFiles)
            If IndexedFiles(j).fileName = selectedFiles(i).fileName And _
               IndexedFiles(j).parentFolderID = selectedFiles(i).parentID Then

                ReDim Preserve result(count)
                result(count) = IndexedFiles(j)
                count = count + 1
                Exit For
            End If
        Next j
    Next i

    GetSelectedIndexedFiles = result
End Function


