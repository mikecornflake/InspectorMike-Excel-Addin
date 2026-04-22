Attribute VB_Name = "libModule"
Option Explicit

Private Sub Test()
    Call ExportModules(Workbooks("InspectorMike_Addin.xlam"), "B:\Code\Office Macros\InspectorMike Excel Addin\Modules")
End Sub

' Error unable to access Visual Basic Project?
' In Excel, goto File -> Options
'    Next: "Trust Center" and click "Trust Center Settings…"
'    Next Select "Macro Settings"
'    Tick "Trust access to the VBA project object model"
Public Sub ExportModules(ByVal AWorkbook As Workbook, ByVal APath As String)
    Dim i As Long
    Dim sName As String
    
    For i = 1 To AWorkbook.VBProject.VBComponents.Count
        Debug.Print AWorkbook.VBProject.VBComponents(i).Name
        
        sName = AWorkbook.VBProject.VBComponents(i).Name
        
        If sName <> "ThisWorkbook" And sName <> "Sheet1" Then
            AWorkbook.VBProject.VBComponents(i).Export (AddTrailingDelimiter(APath) + sName + ".bas")
        End If
    Next i
End Sub
