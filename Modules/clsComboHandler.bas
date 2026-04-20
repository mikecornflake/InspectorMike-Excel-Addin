VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComboHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1

Private mFieldName As String
Private mParentForm As Object   ' late-bound to avoid circular type issues

Public Sub Init(ByVal pCombo As MSForms.ComboBox, _
                ByVal pFieldName As String, _
                ByVal pForm As Object)
    
    Set cbo = pCombo
    mFieldName = pFieldName
    Set mParentForm = pForm
End Sub

Private Sub cbo_Change()
    ' Prevent firing during form load
    If mParentForm.IsLoading Then Exit Sub
    
    mParentForm.RefreshChildLists mFieldName
End Sub

