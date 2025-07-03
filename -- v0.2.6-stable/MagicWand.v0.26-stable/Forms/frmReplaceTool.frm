VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReplaceTool 
   Caption         =   ".:: AFRY BA GOT ::. MagicWand "
   ClientHeight    =   15510
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

Private Sub btnBrowseFolder_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Välj mapp att söka i"
        If .Show = -1 Then
            txtFolderPath.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub btnBrowsePDFPath_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Välj alternativ PDF-mapp"
        If .Show = -1 Then
            txtAltPDFPath.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub btnDateScan_Click()
    ' Rensa tidigare fält
    Dim i As Long
    For i = 1 To 5
        Me.Controls("txtFind" & i).Text = ""
        Me.Controls("txtReplace" & i).Text = ""
        Me.Controls("chkCase" & i).Value = False
        Me.Controls("chkWhole" & i).Value = False
    Next i
    Me.lstSpellingResult.Clear

    ' Kör datumscanning
    Call PreScanIsoDates( _
        folderPath:=Me.txtFolderPath.Text, _
        includeSubfolders:=Me.chkIncludeSubfolders.Value, _
        preserveFolderName:=Me.txtPreserveSubFolder.Text)
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
    ' ?? Rensa alla tidigare sök-/ersättfält och listan
    Dim i As Long
    For i = 1 To 5
        Me.Controls("txtFind" & i).Text = ""
        Me.Controls("txtReplace" & i).Text = ""
        Me.Controls("chkCase" & i).Value = False
        Me.Controls("chkWhole" & i).Value = False
    Next i
    Me.lstSpellingResult.Clear

    ' ?? Kör stavningsanalysen
    Call PreScanSpelling( _
        folderPath:=Me.txtFolderPath.Text, _
        language:=Me.cmbLanguage.Text, _
        includeSubfolders:=Me.chkIncludeSubfolders.Value, _
        preserveFolderName:=Me.txtPreserveSubFolder.Text)
End Sub


Private Sub cmbLanguage_Change()

End Sub

Private Sub lblCreatedBy_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub lblAuthor_Click()
    ' Öppna Teams chatt
    ThisDocument.FollowHyperlink "https://teams.microsoft.com/l/chat/0/0?users=alexander.carlborg@afry.com"
End Sub

Private Sub lblGlobalStats_Click()
Call ShowGlobalEfficiency
End Sub

Private Sub lblLocalStats_Click()
Call ShowLocalEfficiency
End Sub

Private Sub lblProgressBar_Click()

End Sub

Private Sub lblStats_Click()

End Sub

Private Sub lblUser_Click()

End Sub

Private Sub UserForm_Initialize()
    cmbPDFType.AddItem "Normal"
    cmbPDFType.AddItem "PDF/A-1b"
    cmbPDFType.ListIndex = 0

    cmbLanguage.AddItem "Svenska"
    cmbLanguage.AddItem "Engelska"
    cmbLanguage.ListIndex = 0

    chkIncludeSubfolders.Value = True
    chkExportPDF.Value = False
    chkKeepOriginal.Value = False

    lblStatus2.Caption = "Status: Ready"
    lblProgress.Caption = ""
    lblStats.Caption = ""
    lblProgressBar.Width = 0
    
    lblAppVersion.Caption = GetAppVersion()
    Call ShowGlobalEfficiency
    Call ShowLocalEfficiency
           
End Sub


