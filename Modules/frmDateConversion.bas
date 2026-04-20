VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDateConversion 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frmDateConversion.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDateConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnGo_Click()
    Dim sInput As String
    Dim sY As String, sM As String, sd As String
    Dim iY As Long, iM As Long, iD As Long
    Dim i1 As Long, i2 As Long
    
    
    ' 21-1-16
        
    sInput = Trim(edtDate.Value)
    
    i1 = InStr(sInput, "-")
    i2 = InStr(i1 + 1, sInput, "-")
    
    sd = Trim(Left(sInput, i1 - 1))
    sM = Trim(Mid(sInput, i1 + 1, i2 - i1 - 1))
    sY = Trim(Mid(sInput, i2 + 1, 99))
    
    iY = Val(sY)
    iM = Val(sM)
    iD = Val(sd)
    
    If iY < 2000 Then
        iY = iY + 2000
    End If
    
    If iM > 12 Then
        MsgBox "Error"
    End If
    
    sY = Format(iY, "0000")
    sM = Format(iM, "00")
    sd = Format(iD, "00")
    
    edtCorrectDate.Value = sY & "-" & sM & "-" & sd
    edtCorrectDate.Copy
End Sub

Public Sub ShowDateConversionAndApply()
    frmDateConversion.Show
    edtDate.Paste
    btnGo_Click
    Close
End Sub


