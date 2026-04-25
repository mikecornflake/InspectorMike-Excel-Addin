VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXLAdmin 
   Caption         =   "Excel Eventing Administration"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6585
   OleObjectBlob   =   "frmXLAdmin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmXLAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    CentreFormOverExcel Me
End Sub

Private Sub btnValidateForms_Click()
    On Error GoTo CleanExit
    
    btnValidateForms.Enabled = False
    Me.MousePointer = fmMousePointerHourGlass
    Application.Cursor = xlWait
    
    edtResults.Value = ""
    cboMissingTabsheets.Clear
    
    Log "=== Ensuring All Admin Forms exist ==="
    DoEvents
    
    Validate_xe_forms
    Validate_xe_fields
    Validate_xe_lists
    
    Validate_TargetSheets
    
    Log "=== Complete ==="

CleanExit:
    btnValidateForms.Enabled = True
    Me.MousePointer = fmMousePointerDefault
    Application.Cursor = xlDefault
    Log ""
    
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbExclamation, "xlEventing"
    End If
End Sub

Private Sub cmdValidateXEFields_Click()
    On Error GoTo CleanExit
    
    cmdValidateXEFields.Enabled = False
    Me.MousePointer = fmMousePointerHourGlass
    Application.Cursor = xlWait
    
    edtResults.Value = ""
    cboMissingTabsheets.Clear
    
    Log "=== Ensuring existing Workbooks match XE.FIELDS ==="
    DoEvents
    
    ValidateXEFieldsAgainstTargetSheets

CleanExit:
    cmdValidateXEFields.Enabled = True
    Me.MousePointer = fmMousePointerDefault
    Application.Cursor = xlDefault
    Log ""
    
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbExclamation, "xlEventing"
    End If
End Sub

Private Sub Log(ByVal pText As String)
    If Len(edtResults.Value) = 0 Then
        edtResults.Value = pText
    Else
        edtResults.Value = edtResults.Value & vbCrLf & pText
    End If
    
    If Trim(pText) = "" Then
        Application.StatusBar = False
    Else
        Application.StatusBar = pText
    End If
    
    If Application.Cursor = xlWait Then DoEvents
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
    
    ComboBox_RemoveItem cboMissingTabsheets, sheetName
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
        
        Call BasicTidy(ws)
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
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Component", 1, "Installation", "Installation", "combobox", "text", "Y", "", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Component", 2, "Substructure", "Substructure", "combobox", "text", "Y", "", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("Component", 3, "Component", "Component", "combobox", "text", "Y", "", "", ""): r = r + 1
        
        ' GVI
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 1, "Workpack", "Workpack", "combobox", "text", "Y", "WorkpackList", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 2, "Installation", "Installation", "combobox", "text", "Y", "InstallationList", "", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 3, "Substructure", "Substructure", "combobox", "text", "Y", "SubstructureList", "Installation", ""): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 4, "Component", "Component", "combobox", "text", "Y", "ComponentList", "Installation", "Substructure"): r = r + 1
        ws.Cells(r, 1).Resize(1, 10).Value = Array("GVI", 5, "Good_Condition", "Is Component in good condition?", "checkbox", "bool", "Y", "", "", "")
        
        Call BasicTidy(ws)
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
        
        Call BasicTidy(ws)
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
    
    If colFormID <= 0 Or colTargetSheet <= 0 Then Exit Sub
    
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
                    If Not ComboBox_Contains(cboMissingTabsheets, sheetName) Then
                        cboMissingTabsheets.AddItem sheetName
                    End If
                End If
            End If
        End If
    Next iRow
End Sub

Public Sub ValidateXEFieldsAgainstTargetSheets()
    Const SHEET_FIELDS As String = "xe.fields"
    Const SHEET_FORMS As String = "xe.forms"

    Dim wsFields As Worksheet
    Dim wsForms As Worksheet
    Dim dictTargets As Object
    Dim dictExpected As Object
    Dim dictExisting As Object
    Dim dictHandledSheets As Object

    Dim colFormID As Long
    Dim colTargetSheet As Long
    Dim colFieldFormID As Long
    Dim colFieldName As Long
    Dim colDisplayOrder As Long

    Dim lastRow As Long
    Dim iRow As Long

    Dim sFormID As String
    Dim sTargetSheet As String
    Dim sFieldName As String
    Dim wsTarget As Worksheet

    If Not WorksheetExists(SHEET_FIELDS) Then
        Log SHEET_FIELDS & " does not exist."
        Exit Sub
    End If

    If Not WorksheetExists(SHEET_FORMS) Then
        Log SHEET_FORMS & " does not exist."
        If Len(SHEET_FORMS) > 0 Then
            If Not ComboBox_Contains(cboMissingTabsheets, SHEET_FORMS) Then
                cboMissingTabsheets.AddItem SHEET_FORMS
            End If
        End If
        Exit Sub
    End If

    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)
    Set wsForms = ActiveWorkbook.Worksheets(SHEET_FORMS)

    colFormID = FindColumnInSheet(wsForms, "FormID")
    colTargetSheet = FindColumnInSheet(wsForms, "TargetSheet")

    colFieldFormID = FindColumnInSheet(wsFields, "FormID")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    colDisplayOrder = FindColumnInSheet(wsFields, "DisplayOrder")

    If (colFormID <= 0) Or (colTargetSheet <= 0) Or _
       (colFieldFormID <= 0) Or (colFieldName <= 0) Or (colDisplayOrder <= 0) Then

        MsgBox "xe.forms or xe.fields is missing required columns.", vbExclamation, "xlEventing"
        Exit Sub
    End If

    Set dictTargets = CreateObject("Scripting.Dictionary")
    Set dictHandledSheets = CreateObject("Scripting.Dictionary")

    ' Map FormID -> TargetSheet
    lastRow = LastUsedRow(wsForms)

    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsForms.Cells(iRow, colFormID).Value))
        sTargetSheet = Trim$(CStr(wsForms.Cells(iRow, colTargetSheet).Value))

        If Len(sFormID) > 0 And Len(sTargetSheet) > 0 Then
            dictTargets(LCase$(sFormID)) = sTargetSheet
        End If
    Next iRow

    ' Process each FormID once
    lastRow = LastUsedRow(wsFields)

    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsFields.Cells(iRow, colFieldFormID).Value))

        If Len(sFormID) > 0 Then
            If Not dictHandledSheets.Exists(LCase$(sFormID)) Then
                dictHandledSheets.Add LCase$(sFormID), True

                If dictTargets.Exists(LCase$(sFormID)) Then
                    sTargetSheet = CStr(dictTargets(LCase$(sFormID)))

                    If WorksheetExists(sTargetSheet) Then
                        Set wsTarget = ActiveWorkbook.Worksheets(sTargetSheet)
                        ValidateOneXEForm wsFields, wsTarget, sFormID, colFieldFormID, colFieldName, colDisplayOrder
                    Else
                        Log "Target sheet '" & sTargetSheet & "' for FormID '" & sFormID & "' does not exist."
                        If Len(sTargetSheet) > 0 Then
                            If Not ComboBox_Contains(cboMissingTabsheets, sTargetSheet) Then
                                cboMissingTabsheets.AddItem sTargetSheet
                            End If
                        End If
                    End If
                Else
                    Log "FormID '" & sFormID & "' exists in xe.fields but not xe.forms."
                End If
            End If
        End If
    Next iRow

    Log "xe.fields validation against existing tabsheets complete."
End Sub

Private Sub ValidateOneXEForm( _
    ByVal pWSFields As Worksheet, _
    ByVal pWSTarget As Worksheet, _
    ByVal pFormID As String, _
    ByVal pColFormID As Long, _
    ByVal pColFieldName As Long, _
    ByVal pColDisplayOrder As Long)

    Dim expectedFields As Collection
    Dim existingFields As Object
    Dim finalFields As Collection

    Dim lastRow As Long
    Dim lastCol As Long
    Dim iRow As Long
    Dim iCol As Long

    Dim sFieldName As String
    Dim sFormID As String
    Dim vItem As Variant

    Dim arrFields() As String
    Dim arrOrder() As Double
    Dim fieldCount As Long
    Dim j As Long
    Dim tmpField As String
    Dim tmpOrder As Double
    Dim vOrder As Variant

    Dim addedCount As Long
    Dim deletedCount As Long
    Dim movedCount As Long
    Dim expectedCol As Long
    Dim existingCol As Long

    Set expectedFields = New Collection
    Set existingFields = CreateObject("Scripting.Dictionary")
    Set finalFields = New Collection

    Log "Validating FormID '" & pFormID & "' against target sheet '" & pWSTarget.Name & "'..."

    ' Read fields for this FormID
    lastRow = LastUsedRow(pWSFields)

    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(pWSFields.Cells(iRow, pColFormID).Value))

        If StrComp(sFormID, pFormID, vbTextCompare) = 0 Then
            sFieldName = Trim$(CStr(pWSFields.Cells(iRow, pColFieldName).Value))

            If Len(sFieldName) > 0 Then
                fieldCount = fieldCount + 1
                ReDim Preserve arrFields(1 To fieldCount)
                ReDim Preserve arrOrder(1 To fieldCount)

                arrFields(fieldCount) = sFieldName

                vOrder = pWSFields.Cells(iRow, pColDisplayOrder).Value
                If IsNumeric(vOrder) Then
                    arrOrder(fieldCount) = CDbl(vOrder)
                Else
                    arrOrder(fieldCount) = 999999
                    Log "  Warning: field '" & sFieldName & "' has invalid DisplayOrder. Placing at end."
                End If
            End If
        End If
    Next iRow

    If fieldCount = 0 Then
        Log "  No fields defined in xe.fields. Skipping."
        Exit Sub
    End If

    ' Sort by DisplayOrder, then FieldName
    For iRow = 1 To fieldCount - 1
        For j = iRow + 1 To fieldCount
            If (arrOrder(iRow) > arrOrder(j)) Or _
               ((arrOrder(iRow) = arrOrder(j)) And _
                (StrComp(arrFields(iRow), arrFields(j), vbTextCompare) > 0)) Then

                tmpOrder = arrOrder(iRow)
                arrOrder(iRow) = arrOrder(j)
                arrOrder(j) = tmpOrder

                tmpField = arrFields(iRow)
                arrFields(iRow) = arrFields(j)
                arrFields(j) = tmpField
            End If
        Next j
    Next iRow

    ' Build expected field list in sorted order
    For iRow = 1 To fieldCount
        expectedFields.Add arrFields(iRow)
    Next iRow

    ' Existing target headers
    lastCol = LastUsedColumn(pWSTarget)

    For iCol = 1 To lastCol
        sFieldName = Trim$(CStr(pWSTarget.Cells(1, iCol).Value))

        If Len(sFieldName) > 0 Then
            existingFields(LCase$(sFieldName)) = iCol
        End If
    Next iCol

    ' Log missing fields and moved fields
    expectedCol = 0

    For Each vItem In expectedFields
        expectedCol = expectedCol + 1
        sFieldName = CStr(vItem)

        If existingFields.Exists(LCase$(sFieldName)) Then
            existingCol = CLng(existingFields(LCase$(sFieldName)))

            If existingCol <> expectedCol Then
                movedCount = movedCount + 1
                Log "  Move column: '" & sFieldName & "' from column " & existingCol & " to " & expectedCol
            End If
        Else
            addedCount = addedCount + 1
            Log "  Add missing column: '" & sFieldName & "' at column " & expectedCol
        End If
    Next vItem

    ' First: expected fields, in xe.fields order
    For Each vItem In expectedFields
        finalFields.Add CStr(vItem)
    Next vItem

    ' Then: old/deleted/unknown fields
    For iCol = 1 To lastCol
        sFieldName = Trim$(CStr(pWSTarget.Cells(1, iCol).Value))

        If Len(sFieldName) > 0 Then
            If Not CollectionContainsText(expectedFields, sFieldName) Then
                deletedCount = deletedCount + 1
                finalFields.Add sFieldName
                Log "  Move deleted/unknown column to end and mark red: '" & sFieldName & "'"
            End If
        End If
    Next iCol

    If addedCount = 0 And movedCount = 0 And deletedCount = 0 Then
        Log "  No changes required."
    Else
        Log "  Applying changes: " & _
            addedCount & " added, " & _
            movedCount & " moved, " & _
            deletedCount & " marked deleted/unknown."
    End If

    RebuildTargetColumns pWSTarget, expectedFields, finalFields

    On Error Resume Next
    Call BasicTidy(pWSTarget)
    If Err.Number <> 0 Then
        Log "  Warning: BasicTidy failed: " & Err.Description
        Err.Clear
    Else
        Log "  Tidied target sheet."
    End If
    On Error GoTo 0

    Log "Finished validating '" & pFormID & "'."
End Sub
Private Sub RebuildTargetColumns( _
    ByVal pWS As Worksheet, _
    ByVal pExpectedFields As Collection, _
    ByVal pFinalFields As Collection)

    Dim tmpWS As Worksheet
    Dim lastRow As Long
    Dim srcCol As Long
    Dim dstCol As Long
    Dim sFieldName As String
    Dim vItem As Variant
    Dim isDeletedField As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean

    On Error GoTo ErrHandler

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    lastRow = LastUsedRow(pWS)
    If lastRow < 1 Then lastRow = 1

    Set tmpWS = ActiveWorkbook.Worksheets.Add(After:=pWS)
    tmpWS.Name = "__xe_tmp_" & Format$(Now, "hhmmss")

    dstCol = 1

    For Each vItem In pFinalFields
        sFieldName = CStr(vItem)
        srcCol = FindColumnInSheet(pWS, sFieldName)

        If srcCol > 0 Then
            pWS.Columns(srcCol).Copy Destination:=tmpWS.Columns(dstCol)
        Else
            tmpWS.Cells(1, dstCol).Value = sFieldName
        End If

        isDeletedField = Not CollectionContainsText(pExpectedFields, sFieldName)

        If isDeletedField Then
            tmpWS.Columns(dstCol).Font.Color = vbRed
        Else
            tmpWS.Columns(dstCol).Font.Color = vbBlack
        End If

        dstCol = dstCol + 1
    Next vItem

    pWS.Cells.Clear

    tmpWS.Range(tmpWS.Cells(1, 1), tmpWS.Cells(lastRow, pFinalFields.Count)).Copy _
        Destination:=pWS.Cells(1, 1)

    pWS.Activate
    tmpWS.Delete

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Error rebuilding target sheet '" & pWS.Name & "': " & Err.Description, _
           vbExclamation, "xlEventing"
    Resume CleanExit
End Sub

' TODO Refactor use of this with existing Public Function Collection_IndexOf(ByVal pCollection As Collection, ByVal pValue As Variant) As Long
Private Function CollectionContainsText(ByVal pValues As Collection, ByVal pText As String) As Boolean
    Dim vItem As Variant

    For Each vItem In pValues
        If StrComp(CStr(vItem), pText, vbTextCompare) = 0 Then
            CollectionContainsText = True
            Exit Function
        End If
    Next vItem
End Function

