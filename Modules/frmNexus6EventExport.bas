VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNexus6EventExport 
   Caption         =   "Nexus 6 Event Export"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "frmNexus6EventExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNexus6EventExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FInitialised As String

Private Sub btnProcess_Click()
    SaveSettings
    frmNexus6EventExport.Hide
    
    Call ProcessNexus6EventExport("")
End Sub

Private Sub UserForm_Activate()
    If FInitialised = vbNullString Then
        cboServer.AddItem "II5LJSU\SQLEXPRESS2012"
        cboServer.AddItem "INS-SQL-NIC01\SQLEXPRESS"
        cboDatabase.AddItem "Esso_420"
        cboDatabase.AddItem "Esso_UPI"
        cboUser.AddItem "SA"
        cboPassword.AddItem "Net1234."
        
        cboServer.Text = RegistryRead("DOF_Addin", "Nexus 6", "Server", "INS-SQL-NIC01\SQLEXPRESS")
        cboDatabase.Text = RegistryRead("DOF_Addin", "Nexus 6", "Database", "Esso_UPI")
        cboUser.Text = RegistryRead("DOF_Addin", "Nexus 6", "User", "")
        cboPassword.Text = RegistryRead("DOF_Addin", "Nexus 6", "Password", "")
            
        cbPipeline.Value = StringBool(RegistryRead("DOF_Addin", "Nexus 6", "Pipeline", "True"))
        
        FInitialised = "True"
    End If
End Sub

Private Sub SaveSettings()
    Call RegistryWrite("DOF_Addin", "Nexus 6", "Server", Trim(cboServer.Text))
    Call RegistryWrite("DOF_Addin", "Nexus 6", "Database", Trim(cboDatabase.Text))
    Call RegistryWrite("DOF_Addin", "Nexus 6", "User", Trim(cboUser.Text))
    Call RegistryWrite("DOF_Addin", "Nexus 6", "Password", Trim(cboPassword.Text))
    Call RegistryWrite("DOF_Addin", "Nexus 6", "Pipeline", BoolString(cbPipeline.Value))
End Sub

