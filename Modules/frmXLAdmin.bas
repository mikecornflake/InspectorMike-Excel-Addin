VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXLAdmin 
   Caption         =   "Excel Eventing Administration"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmXLAdmin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmXLAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnValidateForms_Click()
    edtResults.Value = ""
    cboMissingTabsheets.Clear
    
    Log "=== xlEventing Admin Validation ==="
    
    Validate_xe_forms
    Validate_xe_fields
    Validate_xe_lists
    
    Validate_TargetSheets
    
    Log "=== Complete ==="
End Sub

Private Sub Log(ByVal pText As String)
    If Len(edtResults.Value) = 0 Then
        edtResults.Value = pText
    Else
        edtResults.Value = edtResults.Value & vbCrLf & pText
    End If
End Sub

Private Sub btnCreateTabsheets_Click()
    Dim sheetName As String
    Dim formID As String
    Dim ws As Worksheet
    
    sheetName = Trim$(cboMissingTabsheets.Value)
    
    If Len(sheetName) = 0 Then
        MsgBox "Please select a missing tabsheet from the list.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    formID = GetFormIDForTargetSheet(sheetName)
    
    If Len(formID) = 0 Then
        MsgBox "No FormID found in xe.forms for target sheet '" & sheetName & "'.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Set ws = EnsureTargetSheetExists(sheetName, formID)
    
    If ws Is Nothing Then
        MsgBox "Failed to create or open sheet '" & sheetName & "'.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Log "Created sheet '" & sheetName & "' from FormID '" & formID & "' with headers only (no data populated)"
    
    RemoveComboItem cboMissingTabsheets, sheetName
End Sub

Private Sub Validate_xe_forms()
    Dim ws As Worksheet
    
    If Not WorksheetExists("xe.forms") Then
        Log "xe.forms tabsheet not found - creating and populating with default data"
        
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "xe.forms"
        ws.Move Before:=ActiveWorkbook.Worksheets(1)
        ColourAdminTab ws
        
        ws.Cells(1, 1).Value = "FormID"
        ws.Cells(1, 2).Value = "Caption"
        ws.Cells(1, 3).Value = "TargetSheet"
        ws.Cells(1, 4).Value = "Type"
        
        ws.Cells(2, 1).Value = "Workpack"
        ws.Cells(2, 2).Value = "Workpack Details"
        ws.Cells(2, 3).Value = "Workpack"
        ws.Cells(2, 4).Value = "Configuration"
        
        ws.Cells(3, 1).Value = "Component"
        ws.Cells(3, 2).Value = "Asset Hierarchy"
        ws.Cells(3, 3).Value = "Component"
        ws.Cells(3, 4).Value = "Configuration"
        
        ws.Cells(4, 1).Value = "GVI"
        ws.Cells(4, 2).Value = "General Visual Inspection"
        ws.Cells(4, 3).Value = "GVI"
        ws.Cells(4, 4).Value = "Event"
        
        BasicTidyWorksheet ws
    Else
        Set ws = ActiveWorkbook.Worksheets("xe.forms")
        Log "xe.forms exists"
        
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            Log "xe.forms was hidden - now visible"
        End If
    End If
End Sub

Private Sub Validate_xe_fields()
    Dim ws As Worksheet
    
    If Not WorksheetExists("xe.fields") Then
        Log "xe.fields tabsheet not found - creating and populating with default data"
        
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "xe.fields"
        ws.Move After:=ActiveWorkbook.Worksheets(2)
        ColourAdminTab ws

        
        ws.Cells(1, 1).Value = "FormID"
        ws.Cells(1, 2).Value = "DisplayOrder"
        ws.Cells(1, 3).Value = "FieldName"
        ws.Cells(1, 4).Value = "Label"
        ws.Cells(1, 5).Value = "ControlType"
        ws.Cells(1, 6).Value = "DataType"
        ws.Cells(1, 7).Value = "Required"
        ws.Cells(1, 8).Value = "ListID"
        ws.Cells(1, 9).Value = "ParentField1"
        ws.Cells(1, 10).Value = "ParentField2"
        
        Dim r As Long: r = 2
        
        ' Workpack
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Workpack", 1, "Name", "Workpack Name", "textbox", "text", "Y", "", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Workpack", 2, "Code", "Workpack Code", "textbox", "text", "N", "", "", ""): r = r + 1
        
        ' Component
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Component", 1, "Installation", "Installation", "combo", "text", "Y", "", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Component", 2, "Substructure", "Substructure", "combo", "text", "Y", "", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Component", 3, "Component", "Component", "combo", "text", "Y", "", "", ""): r = r + 1
        
        ' GVI
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 1, "Workpack", "Workpack", "combo", "text", "Y", "WorkpackList", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 2, "Installation", "Installation", "combo", "text", "Y", "InstallationList", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 3, "Substructure", "Substructure", "combo", "text", "Y", "SubstructureList", "Installation", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 4, "Component", "Component", "combo", "text", "Y", "ComponentList", "Installation", "Substructure"): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 5, "Good_Condition", "Is Component in good condition?", "checkbox", "bool", "Y", "", "", "")
        
        BasicTidyWorksheet ws
    Else
        Set ws = ActiveWorkbook.Worksheets("xe.fields")
        Log "xe.fields exists"
        
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            Log "xe.fields was hidden - now visible"
        End If
    End If
End Sub

Private Sub Validate_xe_lists()
    Dim ws As Worksheet
    
    If Not WorksheetExists("xe.lists") Then
        Log "xe.lists tabsheet not found - creating and populating with default data"
        
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "xe.lists"
        ws.Move After:=ActiveWorkbook.Worksheets(3)
        ColourAdminTab ws

        
        ws.Cells(1, 1).Value = "ListID"
        ws.Cells(1, 2).Value = "SourceSheet"
        ws.Cells(1, 3).Value = "ValueField"
        ws.Cells(1, 4).Value = "FilterField1"
        ws.Cells(1, 5).Value = "FilterParentField1"
        ws.Cells(1, 6).Value = "FilterField2"
        ws.Cells(1, 7).Value = "FilterParentField2"
        ws.Cells(1, 8).Value = "FilterField3"
        ws.Cells(1, 9).Value = "FilterParentField3"
        ws.Cells(1, 10).Value = "DistinctValues"
        ws.Cells(1, 11).Value = "SortValues"
        
        Dim r As Long: r = 2
        
        ws.Cells(r, 1).Resize(1, 11).Value = Array("WorkpackList", "Workpack", "Name", "", "", "", "", "", "", "Y", "Y"): r = r + 1
        ws.Cells(r, 1).Resize(1, 11).Value = Array("InstallationList", "Component", "Installation", "", "", "", "", "", "", "Y", "Y"): r = r + 1
        ws.Cells(r, 1).Resize(1, 11).Value = Array("SubstructureList", "Component", "Substructure", "Installation", "Installation", "", "", "", "", "Y", "Y"): r = r + 1
        ws.Cells(r, 1).Resize(1, 11).Value = Array("ComponentList", "Component", "Component", "Installation", "Installation", "Substructure", "Substructure", "", "", "Y", "Y")
        
        BasicTidyWorksheet ws
    Else
        Set ws = ActiveWorkbook.Worksheets("xe.lists")
        Log "xe.lists exists"
        
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            Log "xe.lists was hidden - now visible"
        End If
    End If
End Sub

Private Sub Validate_TargetSheets()
    Dim wsForms As Worksheet
    Dim lastRow As Long
    Dim iRow As Long
    Dim colFormID As Long
    Dim colTargetSheet As Long
    
    If Not WorksheetExists("xe.forms") Then Exit Sub
    
    Set wsForms = ActiveWorkbook.Worksheets("xe.forms")
    
    colFormID = FindColumnInSheet(wsForms, "FormID")
    colTargetSheet = FindColumnInSheet(wsForms, "TargetSheet")
    
    If colFormID = 0 Or colTargetSheet = 0 Then Exit Sub
    
    lastRow = LastUsedRow(wsForms)
    
    For iRow = 2 To lastRow
        Dim formID As String
        Dim sheetName As String
        
        formID = Trim$(CStr(wsForms.Cells(iRow, colFormID).Value))
        sheetName = Trim$(CStr(wsForms.Cells(iRow, colTargetSheet).Value))
        
        If Len(sheetName) > 0 Then
            If WorksheetExists(sheetName) Then
                Log formID & ": sheet '" & sheetName & "' exists"
                
                Dim ws As Worksheet
                Set ws = ActiveWorkbook.Worksheets(sheetName)
                
                If ws.Visible <> xlSheetVisible Then
                    ws.Visible = xlSheetVisible
                    Log "  -> was hidden, now visible"
                End If
            Else
                Log formID & ": sheet '" & sheetName & "' MISSING"

                If Len(sheetName) > 0 Then
                    If Not ComboContains(cboMissingTabsheets, sheetName) Then
                        cboMissingTabsheets.AddItem sheetName
                    End If
                End If
            End If
        End If
    Next iRow
End Sub

