Attribute VB_Name = "libHTML"
Option Explicit
Option Private Module

Public Sub ShowHelp(ByVal page As String)
    Dim path As String
    
    path = ThisWorkbook.path & "\InspectorMike_Addin_docs\" & page
    
    If Dir(path) <> "" Then
        ThisWorkbook.FollowHyperlink path
    Else
        MsgBox "Help page not found: " & page, vbExclamation
    End If
End Sub
