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
        .Title = "Välj alternativ PDF-mapp"
        If .Show = -1 Then
            txtAltPDFPath.text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub btnStart_Click()
    If Not ValidateProcessingSettings() Then Exit Sub
    StartProcessing
End Sub

Private Sub StartProcessing()
    Dim findTexts(1 To 5) As String
    Dim replaceTexts(1 To 5) As String
    Dim caseFlags(1 To 5) As Boolean
    Dim wordFlags(1 To 5) As Boolean
    Dim i As Long

    For i = 1 To 5
        findTexts(i) = Me.Controls("txtFind" & i).text
        replaceTexts(i) = Me.Controls("txtReplace" & i).text
        caseFlags(i) = Me.Controls("chkCase" & i).Value
        wordFlags(i) = Me.Controls("chkWhole" & i).Value
    Next i

    ProcessIndexedDocuments findTexts, replaceTexts, caseFlags, wordFlags, _
        Me.chkExportPDF.Value, Me.cmbPDFType.text, Me.txtAltPDFPath.text, _
        Me.txtPrefix.text, Me.txtSuffix.text, Me.chkKeepOriginal.Value, _
        Me.cmbLanguage.text
End Sub

Private Sub btnCancel_Click()
    cancelRequested = True
    lblStatus2.Caption = "Status: Aborting..."
End Sub

Private Sub btnSpellcheck_Click()
    ' Rensa alla tidigare sök-/ersättfält och listan
    Dim i As Long
    For i = 1 To 5
        Me.Controls("txtFind" & i).text = ""
        Me.Controls("txtReplace" & i).text = ""
        Me.Controls("chkCase" & i).Value = False
        Me.Controls("chkWhole" & i).Value = False
    Next i
    Me.lstSpellingResult.Clear

' === Kör stavningsanalysen på indexerade filer ===
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
    ' Öppna Teams chatt
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

    ' === Läs in sparade filer från frmIndexBrowser ===
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









