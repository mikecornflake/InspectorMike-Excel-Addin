Attribute VB_Name = "libRegistry"
Option Explicit


'Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings

Public Sub RegistryWrite(AApp As String, ASection As String, AKey As String, AValue As Variant)
    SaveSetting AApp, ASection, AKey, AValue
End Sub

Public Function RegistryRead(AApp As String, ASection As String, AKey As String, AValue As Variant) As String
    RegistryRead = GetSetting(AApp, ASection, AKey, AValue)
End Function

Public Sub RegistryDelete(AApp As String, ASection As String, Optional AKey As String = "")
    If AKey = "" Then
        DeleteSetting AApp, ASection
    Else
        DeleteSetting AApp, ASection, AKey
    End If
End Sub




