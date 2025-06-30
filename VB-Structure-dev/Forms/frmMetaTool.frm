VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetaTool 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMetaTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMetaTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = ":: MagicWand | Metadata Tool"

    With cmbMetaFieldGlobal
        .AddItem "Title"
        .AddItem "Subject"
        .AddItem "Author"
        .AddItem "Keywords"
    End With

    With cmbSourceGlobal
        .AddItem "Static Text"
        .AddItem "Current User"
        .AddItem "Filename"
        .AddItem "Last Saved Date"
        .AddItem "Current Date"
    End With
End Sub

Private Sub btnApplyGlobal_Click()
    ApplyGlobalMetadata cmbMetaFieldGlobal.text, cmbSourceGlobal.text, txtValueGlobal.text, chkOverwriteEmptyOnly.Value
    MsgBox "Global metadata applied.", vbInformation
End Sub

Private Sub btnAddRule_Click()
    ' TODO: Lägg till en ny rad i lstMetaMatrix
End Sub

Private Sub btnRemoveRule_Click()
    ' TODO: Ta bort markerad rad
End Sub

Private Sub btnApplyMatrix_Click()
    ' TODO: Läs alla regler och kör ApplyMetadataMatrix
End Sub

