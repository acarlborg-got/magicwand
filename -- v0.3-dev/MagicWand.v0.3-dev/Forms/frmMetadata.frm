VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetadata 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMetadata.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()
Option Explicit

Private Sub btnInject_Click()
    Dim meta As FileMetadata
    meta.Title = txtTitle.text
    meta.Subject = txtSubject.text
    meta.Author = txtAuthor.text
    meta.Keywords = txtKeywords.text
    meta.DocumentDate = txtDate.text

    Dim files() As IndexedFile
    files = GetSelectedIndexedFiles()
    Dim i As Long
    For i = 0 To UBound(files)
        Call WriteMetadata(files(i).filePath, meta)
    Next i
    MsgBox "Metadata injected to selected files.", vbInformation
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = ":: MagicWand | Metadata"
    txtTitle.text = ""
    txtSubject.text = ""
    txtAuthor.text = ""
    txtKeywords.text = ""
    txtDate.text = Format(Date, "yyyy-mm-dd")
End Sub

