VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompareSheets 
   Caption         =   "Compare Worksheets"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "frmCompareSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCompareSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Init()
    Dim oSheet As Worksheet
    
    cboSheet1.Clear
    cboSheet2.Clear
    edtQCSheet.Text = "Comparison"
    
    For Each oSheet In ActiveWorkbook.Worksheets
        cboSheet1.AddItem oSheet.Name
        cboSheet2.AddItem oSheet.Name
    Next oSheet
    
    cboSheet1.Value = cboSheet1.List(1)
    cboSheet2.Value = cboSheet1.List(2)
    
    edtQCSheet.Text = cboSheet1.Value + " QC"
End Sub


Private Sub btnCompare_Click()
    Call CompareSheets(cboSheet1.Text, cboSheet2.Text, edtQCSheet.Text)
    
    Hide
End Sub


Private Sub cboSheet1_Change()
  edtQCSheet.Text = cboSheet1.Value + " QC"
End Sub

Private Sub UserForm_Activate()
    Call Init
End Sub


