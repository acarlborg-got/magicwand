VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIndexBrowser 
   Caption         =   ".:: AFRY BA GOT ::. MagicWand  | Index folders"
   ClientHeight    =   10530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14820
   OleObjectBlob   =   "frmIndexBrowser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIndexBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



' === Modulnivåtyper ===
Private DisplayedFolderIDs() As Long

Private Sub cmbNextForm_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblAuthor_Click()
    ' Öppna Teams chatt
    ThisDocument.FollowHyperlink "https://teams.microsoft.com/l/chat/0/0?users=alexander.carlborg@afry.com"
End Sub

Private Sub txtRootFolder_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = GetTitle_frmIndexBrowser()
    lblAppVersion.Caption = GetAppVersion()

    ' Initiera nästa-formular lista
    With cmbNextForm
        .AddItem "Search & Replace"
        .AddItem "Metadata injection"
        .AddItem "Format cleanup"
    End With
    ' UI-kontroller
    btnGo.Enabled = False
    btnGo.BackColor = vbWhite
End Sub


' === Browse-klick ===
Private Sub btnBrowseFolder_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then txtRootFolder.text = .SelectedItems(1)
    End With
End Sub

' === Index-klick ===
Private Sub btnLoad_Click()
    If txtRootFolder.text = "" Then Exit Sub
    Call IndexFoldersAndFiles(txtRootFolder.text)
    Call BuildFolderList
    Call UpdateFilesFromSelectedFolders
    Call UpdateStats
End Sub

' === Lista mappar med indrag ===
Private Sub BuildFolderList()
    Dim i As Long, indent As String
    lstFolders.Clear
    ReDim DisplayedFolderIDs(UBound(IndexedFolders))

    For i = 0 To UBound(IndexedFolders)
        Select Case IndexedFolders(i).depth
            Case 0: indent = ""
            Case 1: indent = "- "
            Case Else: indent = String((IndexedFolders(i).depth - 1) * 2, " ") & "- "
        End Select

        lstFolders.AddItem indent & GetIndexedFolderName(IndexedFolders(i).folderPath)
        DisplayedFolderIDs(i) = IndexedFolders(i).ID
    Next i

    lstFolders.MultiSelect = fmMultiSelectMulti
End Sub

Private Sub lblAppVersion_Click()

End Sub

Private Sub lblStats_Click()

End Sub

' === Ändring i mappval ===
Private Sub lstFolders_Change()
    Dim i As Long
    For i = 0 To lstFolders.ListCount - 1
        IndexedFolders(DisplayedFolderIDs(i)).selected = lstFolders.selected(i)
    Next i

    Call UpdateFilesFromSelectedFolders
    Call UpdateStats
End Sub

' === Visa filer från valda mappar ===
Private Sub UpdateFilesFromSelectedFolders()
    Dim i As Long, j As Long
    listFiles.Clear

    For i = 0 To UBound(IndexedFiles)
        IndexedFiles(i).selected = False
    Next i

    For i = 0 To UBound(IndexedFolders)
        If IndexedFolders(i).selected Then
            For j = 0 To UBound(IndexedFiles)
                If IndexedFiles(j).parentFolderID = IndexedFolders(i).ID Then
                    If Not FileAlreadyListed(listFiles, IndexedFiles(j).fileName) Then
                        listFiles.AddItem IndexedFiles(j).fileName
                        listFiles.selected(listFiles.ListCount - 1) = True
                        IndexedFiles(j).selected = True
                    End If
                End If
            Next j
        End If
    Next i
End Sub

' === Filtrera dubbletter ===
Private Function FileAlreadyListed(lst As MSForms.ListBox, fileName As String) As Boolean
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.List(i) = fileName Then
            FileAlreadyListed = True
            Exit Function
        End If
    Next i
    FileAlreadyListed = False
End Function

' === Ändra valtillstånd i IndexedFiles ===
Private Sub listFiles_Change()
    Dim i As Long, j As Long, fname As String

    For i = 0 To listFiles.ListCount - 1
        fname = listFiles.List(i)
        For j = 0 To UBound(IndexedFiles)
            If IndexedFiles(j).fileName = fname Then
                IndexedFiles(j).selected = listFiles.selected(i)
                Exit For
            End If
        Next j
    Next i

    Call UpdateStats
End Sub

' === Markera fil som vald manuellt ===
Private Sub SetFileSelectedStatus(fileName As String, sel As Boolean)
    Dim i As Long
    For i = 0 To UBound(IndexedFiles)
        If IndexedFiles(i).fileName = fileName Then
            IndexedFiles(i).selected = sel
            Exit Sub
        End If
    Next i
End Sub

' === Lägga till valda filer till lstFilesToProcess ===
Private Sub btnAddFiles_Click()
    Dim i As Long, fname As String

    For i = 0 To listFiles.ListCount - 1
        If listFiles.selected(i) Then
            fname = listFiles.List(i)
            If Not FileAlreadyListed(lstFilesToProcess, fname) Then
                lstFilesToProcess.AddItem fname
                Call SetFileSelectedStatus(fname, True)
            End If
        End If
    Next i

    Call UpdateStats
    Call ResetSaveUI

End Sub

' === Ta bort valda filer från lstFilesToProcess ===
Private Sub btnRemFiles_Click()
    Dim i As Long
    For i = lstFilesToProcess.ListCount - 1 To 0 Step -1
        If lstFilesToProcess.selected(i) Then
            Call SetFileSelectedStatus(lstFilesToProcess.List(i), False)
            lstFilesToProcess.RemoveItem i
        End If
    Next i

    Call UpdateStats
    Call ResetSaveUI
End Sub

' === Spara urval till arrays ===
Public Sub btnSaveSelection_Click()
    Dim i As Long, j As Long
    Dim fCount As Long, dCount As Long
    Dim fname As String
    Dim folderIDsDict As Object
    Set folderIDsDict = CreateObject("Scripting.Dictionary")

    Erase selectedFiles
    Erase selectedFolders

    If lstFilesToProcess.ListCount = 0 Then
        MsgBox "No files selected to save.", vbExclamation
        Exit Sub
    End If

    For i = 0 To lstFilesToProcess.ListCount - 1
        fname = lstFilesToProcess.List(i)
        For j = 0 To UBound(IndexedFiles)
            If IndexedFiles(j).fileName = fname Then
                ReDim Preserve selectedFiles(fCount)
                selectedFiles(fCount).fileName = IndexedFiles(j).fileName
                selectedFiles(fCount).parentID = IndexedFiles(j).parentFolderID
                selectedFiles(fCount).filePath = IndexedFiles(j).filePath
                fCount = fCount + 1

                If Not folderIDsDict.Exists(CStr(IndexedFiles(j).parentFolderID)) Then
                    folderIDsDict.Add CStr(IndexedFiles(j).parentFolderID), IndexedFiles(j).parentFolderID
                End If
                Exit For
            End If
        Next j
    Next i

    Dim key As Variant
    For Each key In folderIDsDict.keys
        ReDim Preserve selectedFolders(dCount)
        selectedFolders(dCount) = IndexedFolders(folderIDsDict(key))
        dCount = dCount + 1
    Next key

    ' Visuell feedback
    lstFilesToProcess.BackColor = RGB(240, 240, 240)
    BtnSaveSelection.BackColor = RGB(200, 200, 200)
    BtnSaveSelection.Caption = "Selection Saved"

    btnGo.BackColor = RGB(200, 255, 200)
    btnGo.Enabled = True
    cmbNextForm.BackColor = vbWhite

    ' MsgBox fCount & " files saved to selection.", vbInformation
End Sub

Private Sub ResetSaveUI()
    lstFilesToProcess.BackColor = vbWhite
    BtnSaveSelection.BackColor = RGB(200, 255, 200)   ' Grön (spara krävs)
    BtnSaveSelection.Caption = "Save Selection"
    
    btnGo.BackColor = vbWhite
    btnGo.Enabled = False                             ' Blockerad
    cmbNextForm.BackColor = vbWhite                   ' Alltid vit
End Sub


Private Sub lstFilesToProcess_Change()
    Call ResetSaveUI
    Call UpdateStats
End Sub


' === Statistikruta längst ner ===
Private Sub UpdateStats()
    Dim fCount As Long, dCount As Long
    Dim i As Long, j As Long
    Dim tempFolderDict As Object
    Set tempFolderDict = CreateObject("Scripting.Dictionary")

    ' Räkna endast från "Files to Process"
    fCount = lstFilesToProcess.ListCount

    For i = 0 To lstFilesToProcess.ListCount - 1
        Dim fname As String
        fname = lstFilesToProcess.List(i)

        ' Hitta parentFolderID för varje fil
        For j = 0 To UBound(IndexedFiles)
            If IndexedFiles(j).fileName = fname Then
                If Not tempFolderDict.Exists(IndexedFiles(j).parentFolderID) Then
                    tempFolderDict.Add IndexedFiles(j).parentFolderID, True
                End If
                Exit For
            End If
        Next j
    Next i

    dCount = tempFolderDict.count

    lblStats.Caption = fCount & " file(s) selected from " & dCount & " folder(s)."
End Sub

Private Sub btnGo_Click()
    If lstFilesToProcess.ListCount = 0 Then
        MsgBox "No files selected to process.", vbExclamation
        Exit Sub
    End If

    ' Spara det aktuella urvalet
    Call btnSaveSelection_Click

    ' Öppna rätt formulär
    Select Case cmbNextForm.Value
        Case "Search & Replace"
            ShowFormSafe "frmReplaceTool"

        Case "Metadata injection"
            ShowFormSafe "frmMetadata"

        Case "Format cleanup"
            ShowFormSafe "frmFormatter"

        Case Else
            ' MsgBox "Please select a module to continue.", vbInformation
    End Select
End Sub

Private Sub UserForm_Click()
End Sub



