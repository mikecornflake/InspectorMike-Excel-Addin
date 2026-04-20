VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptions 
   Caption         =   "Options"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6225
   OleObjectBlob   =   "frmOptions.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    lbEventCodes.AddItem ("ST.RKB")
    ' lbEventCodes.AddItem ("SP.SPS")
    ' lbEventCodes.AddItem ("SP.SPE")
    
    lbIncidentTypes.AddItem ("FJ")
    lbIncidentTypes.AddItem ("AN")
    
    RefreshUI
End Sub

Private Sub RefreshUI()
    If FAscendingInspection = vbNull Then
        FAscendingInspection = True
    End If
    
    optAsc.Value = FAscendingInspection
    optDesc.Value = Not FAscendingInspection
    
    edtKP.Value = FKPThresholdForSameness
    
    btnTidyListing.Enabled = Not mod_VisualSoft.IsTidy
    btnTidyAnomaly.Enabled = mod_VisualSoft.IsTidy
    btnQCChecks.Enabled = mod_VisualSoft.IsTidy And Not mod_VisualSoft.IsQC
    btnUpdateKPLength.Enabled = mod_VisualSoft.IsTidy
    btnDMWTCHack.Enabled = mod_VisualSoft.IsTidy
    btnSetInspectionEndPosToNextInspectionStart.Enabled = mod_VisualSoft.IsQC
    
    btnAddEventCode.Enabled = Trim(edtEventCode.Text) <> ""
    btnAddIncidentType.Enabled = Trim(edtIncidentType.Text) <> ""
    
    btnDeleteEventCode.Enabled = lbEventCodes.ListIndex <> -1
    btnDeleteIncidentType.Enabled = lbIncidentTypes.ListIndex <> -1
End Sub

Private Sub btnTidyListing_Click()
    mod_VisualSoft.TidyVWExcelExport
    
    RefreshUI
End Sub

Private Sub btnTidyAnomaly_Click()
    mod_VisualSoft.ProcessAnomalies
    
    RefreshUI
    Hide
End Sub

Private Sub btnSetInspectionEndPosToNextInspectionStart_Click()
    mod_VisualSoft.SetInspectionEndPosToNextInspectionStart
    
    RefreshUI
End Sub

Private Sub btnUpdateKPLength_Click()
    mod_VisualSoft.UpdateKPLength
    
    RefreshUI
End Sub

Private Sub btnDMWTCHack_Click()
    mod_VisualSoft.ApplyDM_WTC_Hack
    
    RefreshUI
End Sub

Private Sub btnQCChecks_Click()
    Dim i As Integer
    
    ' Populate the Arrays used in the QC Checks
    ReDim FDuplicateEventCodeChecks(lbEventCodes.ListCount)
    For i = 0 To lbEventCodes.ListCount - 1
        FDuplicateEventCodeChecks(i).Code = lbEventCodes.List(i)
        FDuplicateEventCodeChecks(i).lastRow = -1
    Next i
    
    ReDim FDuplicateIncidentTypeChecks(lbIncidentTypes.ListCount)
    For i = 0 To lbIncidentTypes.ListCount - 1
        FDuplicateIncidentTypeChecks(i).Code = lbIncidentTypes.List(i)
        FDuplicateIncidentTypeChecks(i).lastRow = -1
    Next i
    
    mod_VisualSoft.QCChecks
    RefreshUI
    Hide
End Sub

Private Sub btnAddIncidentType_Click()
    lbIncidentTypes.AddItem (edtIncidentType.Text)
    edtIncidentType.Text = ""
    RefreshUI
End Sub

Private Sub btnAddEventCode_Click()
    lbEventCodes.AddItem (edtEventCode.Text)
    edtEventCode.Text = ""
    RefreshUI
End Sub

Private Sub btnDeleteEventCode_Click()
    If lbEventCodes.ListIndex <> -1 Then
        lbEventCodes.RemoveItem lbEventCodes.ListIndex
    End If
    RefreshUI
End Sub

Private Sub btnDeleteIncidentType_Click()
    If lbIncidentTypes.ListIndex <> -1 Then
        lbIncidentTypes.RemoveItem lbIncidentTypes.ListIndex
    End If
    RefreshUI
End Sub

Private Sub edtEventCode_Change()
    btnAddEventCode.Enabled = Trim(edtEventCode.Text) <> ""
End Sub

Private Sub edtIncidentType_Change()
    btnAddIncidentType.Enabled = Trim(edtIncidentType.Text) <> ""
End Sub

Private Sub lbEventCodes_Change()
    btnDeleteEventCode.Enabled = lbEventCodes.ListIndex <> -1
End Sub

Private Sub lbIncidentTypes_Change()
    btnDeleteIncidentType.Enabled = lbIncidentTypes.ListIndex <> -1
End Sub

Private Sub edtKP_Change()
    FKPThresholdForSameness = edtKP.Value
End Sub

Private Sub optAsc_Change()
    FAscendingInspection = optAsc.Value
End Sub

Private Sub optDesc_Change()
    FAscendingInspection = optAsc.Value
End Sub


