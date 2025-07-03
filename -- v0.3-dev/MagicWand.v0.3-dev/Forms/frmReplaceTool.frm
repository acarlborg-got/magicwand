VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReplaceTool 
   Caption         =   ".:: AFRY BA GOT ::. MagicWand "
   ClientHeight    =   14010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13455
   OleObjectBlob   =   "frmReplaceTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReplaceTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBrowsePDFPath_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "V�lj alternativ PDF-mapp"
        If .Show = -1 Then
            txtAltPDFPath.text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub btnStart_Click()
    If Not ValidateProcessingSettings() Then Exit Sub
    StartProcessing
End Sub

Private Sub btnCancel_Click()
    cancelRequested = True
    lblStatus2.Caption = "Status: Aborting..."
End Sub

Private Sub btnSpellcheck_Click()
    ' Rensa alla tidigare s�k-/ers�ttf�lt och listan
    Dim i As Long
    For i = 1 To 5
        Controls("txtFind" & i).text = ""
        Controls("txtReplace" & i).text = ""
        Controls("chkCase" & i).Value = False
        Controls("chkWhole" & i).Value = False
    Next i
    Me.lstSpellingResult.Clear

' === K�r stavningsanalysen p� indexerade filer ===
Dim filePaths() As String
filePaths = GetSelectedFilePaths()

Call PreScanSpelling( _
    filePaths:=filePaths, _
    language:=Me.cmbLanguage.text, _
    preserveFolderName:=Me.txtPreserveSubFolder.text)

End Sub
Private Sub btnFindDates_Click()
    Call FindDatesInIndexedFiles
End Sub
Private Sub cmbLanguage_Change()

End Sub

Private Sub lblCreatedBy_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub lblAppVersion_Click()

End Sub

Private Sub lblAuthor_Click()
    ' �ppna Teams chatt
    ThisDocument.FollowHyperlink "https://teams.microsoft.com/l/chat/0/0?users=alexander.carlborg@afry.com"
End Sub

Private Sub lblProgressBar_Click()

End Sub

Private Sub lblStats_Click()

End Sub

Private Sub lstSpellingResult_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = GetTitle_frmReplaceTool()
    lblAppVersion.Caption = GetAppVersion()

    cmbPDFType.Clear
    cmbPDFType.AddItem "Normal"
    cmbPDFType.AddItem "PDF/A-1b"
    cmbPDFType.ListIndex = 0

    cmbLanguage.Clear
    cmbLanguage.AddItem "Svenska"
    cmbLanguage.AddItem "Engelska"
    cmbLanguage.ListIndex = 0

    chkExportPDF.Value = False
    chkKeepOriginal.Value = False

    lblStatus2.Caption = "Status: Ready"
    lblProgress.Caption = ""
    lblStats.Caption = ""
    lblProgressBar.Width = 0

    ' === L�s in sparade filer fr�n frmIndexBrowser ===
    Dim files() As tFileSelection
    Dim fileCount As Long

    On Error GoTo NoFiles
    files = GetSelectedFiles()

    If (Not Not files) = False Or UBound(files) < 0 Then GoTo NoFiles

    fileCount = UBound(files) + 1
    lblStats.Caption = "Loaded: " & fileCount & " file(s) from saved index"
    Exit Sub

NoFiles:
    lblStats.Caption = "No files loaded from index."
End Sub









