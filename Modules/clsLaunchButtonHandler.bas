VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLaunchButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private mFormID As String
Private mParentForm As Object

Public Sub Init(ByVal pButton As MSForms.CommandButton, _
                ByVal pFormID As String, _
                ByVal pForm As Object)
    Set btn = pButton
    mFormID = pFormID
    Set mParentForm = pForm
End Sub

Private Sub btn_Click()
    mParentForm.LaunchForm mFormID
End Sub

