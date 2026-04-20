VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEventMerge 
   Caption         =   "Event Merge"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   OleObjectBlob   =   "frmEventMerge.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEventMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    cboCode.AddItem "AN"
    cboCode.AddItem "FJ"
    
    cboMatchCode.AddItem "1"
    cboMatchCode.AddItem "0"
    cboMatchCode.AddItem "-1"
    
    edtBaseFolder.Text = FBaseFolder
    edtCurrent.Text = fCurrent
    edtTrack.Text = FTrack
    edtHistorical.Text = FHistorical
    cboCode.Text = FCode
    cboMatchCode.Text = FMatchCode
End Sub


Private Sub btnGo_Click()
    FBaseFolder = edtBaseFolder.Text
    fCurrent = edtCurrent.Text
    FTrack = edtTrack.Text
    FHistorical = edtHistorical.Text
    FCode = cboCode.Text
    FMatchCode = cboMatchCode.Text
    
    MergeEvents
End Sub




