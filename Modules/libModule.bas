Attribute VB_Name = "libModule"
Option Explicit
Option Private Module

Private Sub Test()
    Call ExportModules(Workbooks("InspectorMike_Addin.xlam"), "B:\Code\Office Macros\InspectorMike Excel Addin\Modules")
End Sub

' Error unable to access Visual Basic Project?
' In Excel, goto File -> Options
'    Next: "Trust Center" and click "Trust Center Settings…"
'    Next Select "Macro Settings"
'    Tick "Trust access to the VBA project object model"
Private Sub ExportModules(ByVal AWorkbook As Workbook, ByVal APath As String)
    Dim i As Long
    Dim sName As String
    
    If AWorkbook Is Nothing Then Err.Raise vbObjectError + 2, , "Workbook is Nothing"
    
    If AWorkbook.path = "" Then Err.Raise vbObjectError + 3, , "Workbook must be saved first"

    ' Save workbook first
    If Not AWorkbook.Saved Then
        Application.DisplayAlerts = False
        AWorkbook.Save
        Application.DisplayAlerts = True
    End If

    For i = 1 To AWorkbook.VBProject.VBComponents.Count
        Debug.Print AWorkbook.VBProject.VBComponents(i).Name
        
        sName = AWorkbook.VBProject.VBComponents(i).Name
        
        If sName <> "ThisWorkbook" And sName <> "Sheet1" Then
            AWorkbook.VBProject.VBComponents(i).Export _
                Path_AddTrailingDelimiter(APath) & sName & ".bas"
        End If
    Next i
End Sub
