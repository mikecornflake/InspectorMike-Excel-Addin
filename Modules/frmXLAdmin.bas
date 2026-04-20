VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXLAdmin 
   Caption         =   "Excel Eventing Administration"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmXLAdmin.frx":0000
   StartUpPosition =   1  'CenterOwner
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
    
    sheetName = Trim$(cboMissingTabsheets.Value)
    
    If Len(sheetName) = 0 Then
        MsgBox "Please select a missing tabsheet from the list.", vbExclamation
        Exit Sub
    End If
    
    If WorksheetExists(sheetName) Then
        Log "Sheet already exists!"
        Exit Sub
    End If
    
    CreateTargetSheetFromFields sheetName
End Sub

Private Sub Validate_xe_forms()
    Dim ws As Worksheet
    
    If Not WorksheetExists("xe.forms") Then
        Log "xe.forms tabsheet not found - creating and populating with default data"
        
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "xe.forms"
        
        ws.Cells(1, 1).Value = "FormID"
        ws.Cells(1, 2).Value = "Caption"
        ws.Cells(1, 3).Value = "TargetSheet"
        
        ws.Cells(2, 1).Value = "Workpack"
        ws.Cells(2, 2).Value = "Workpack Details"
        ws.Cells(2, 3).Value = "Workpack"
        
        ws.Cells(3, 1).Value = "Component"
        ws.Cells(3, 2).Value = "Asset Hierarchy"
        ws.Cells(3, 3).Value = "Component"
        
        ws.Cells(4, 1).Value = "GVI"
        ws.Cells(4, 2).Value = "General Visual Inspection"
        ws.Cells(4, 3).Value = "GVI"
        
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
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 5, "Good Condition?", "Is the Component in acceptable condition?", "textbox", "text", "Y", "", "", "")
        
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
        
    Else
        Set ws = ActiveWorkbook.Worksheets("xe.lists")
        Log "xe.lists exists"
        
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            Log "xe.lists was hidden - now visible"
        End If
    End If
End Sub

Private Function ComboContains(ByVal cbo As MSForms.ComboBox, ByVal pValue As String) As Boolean
    Dim i As Long
    
    ComboContains = False
    
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), pValue, vbTextCompare) = 0 Then
            ComboContains = True
            Exit Function
        End If
    Next i
End Function

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

Private Sub CreateTargetSheetFromFields(ByVal pSheetName As String)
    Const SHEET_FIELDS As String = "xe.fields"
    
    Dim wsNew As Worksheet
    Dim wsFields As Worksheet
    
    Dim colFormID As Long
    Dim colFieldName As Long
    Dim colDisplayOrder As Long
    
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim fields As Collection
    Dim item As Variant
    
    If Not WorksheetExists(SHEET_FIELDS) Then
        Log "xe.fields not found!"
        Exit Sub
    End If
    
    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)
    
    colFormID = FindColumnInSheet(wsFields, "FormID")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    colDisplayOrder = FindColumnInSheet(wsFields, "DisplayOrder")
    
    If colFormID = 0 Or colFieldName = 0 Or colDisplayOrder = 0 Then
        Log "xe.fields missing required columns!"
        Exit Sub
    End If
    
    ' Gather fields for this form
    Set fields = New Collection
    
    lastRow = LastUsedRow(wsFields)
    
    For iRow = 2 To lastRow
        If StrComp(Trim$(wsFields.Cells(iRow, colFormID).Value), pSheetName, vbTextCompare) = 0 Then
            fields.Add Array( _
                wsFields.Cells(iRow, colDisplayOrder).Value, _
                wsFields.Cells(iRow, colFieldName).Value _
            )
        End If
    Next iRow
    
    If fields.Count = 0 Then
        Log "No field definitions found for '" & pSheetName & "'!"
        Exit Sub
    End If
    
    ' Sort by DisplayOrder
    Set fields = CollectionBubbleSort(fields)
    
    ' Create sheet
    Set wsNew = Add_Sheet(pSheetName, Sheets.Count + 1)
    
    ' Write headers
    Dim col As Long: col = 1
    
    For Each item In fields
        wsNew.Cells(1, col).Value = item(1) ' FieldName
        col = col + 1
    Next item
    
    ' Format header row (optional but nice)
    wsNew.Rows(1).Font.Bold = True
    
    Call BasicTidyWorksheet(wsNew)
    
    Log "Created sheet '" & pSheetName & "' with headers only (no data populated)"
End Sub

