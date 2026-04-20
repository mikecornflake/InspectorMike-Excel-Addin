VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRenameColumns 
   Caption         =   "Rename Columns"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmRenameColumns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRenameColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnRenameNewToOriginal_Click()
    RenameColumns (False)
End Sub

Private Sub btnRenameOriginalToNew_Click()
    RenameColumns (True)
End Sub

Private Sub btnTabsheet_Click()
    AddColumnNamesLookup
    
    frmRenameColumns.Hide
End Sub
