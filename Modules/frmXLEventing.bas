VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXLEventing 
   Caption         =   "frmXLEventing"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   OleObjectBlob   =   "frmXLEventing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmXLEventing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFormID As String 'xe.forms "FormID"
Private mActiveRow As Long ' Current Row in the "TargetSheet" for active "FormID"
Private mFieldDefs As Collection
Private mControlMap As Object ' Links a dynamically created control for each fieldname (which are found in xe.lists)
Private mEventHandlers As Collection ' EventHandlers for ComboBoxes so related ComboBoxes can be refreshed
Private mIsLoading As Boolean

Public Sub SetupForm(ByVal pFormID As String, ByVal pActiveRow As Long)
    mIsLoading = True
    On Error GoTo ErrHandler
    
    mFormID = pFormID
    mActiveRow = pActiveRow
    
    Set mEventHandlers = New Collection

    ClearDynamicControls
    BuildControls
    PopulateAllLists
    
    If mActiveRow >= 2 Then
        LoadRowValues False
        PopulateAllLists
        LoadRowValues True
    End If
    
    LayoutForm

CleanExit:
    mIsLoading = False
    Exit Sub

ErrHandler:
    MsgBox "Error in SetupForm: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveFormToRow
End Sub

Public Function IsLoading() As Boolean
    IsLoading = mIsLoading
End Function

Private Sub UserForm_Resize()
    LayoutForm
End Sub

Private Sub LayoutForm()
    Dim margin As Single
    Dim buttonTop As Single

    margin = 12

    lblTitle.Left = margin
    lblTitle.Top = margin
    lblTitle.Width = Me.InsideWidth - margin * 2

    buttonTop = Me.InsideHeight - 36 - margin

    fraHost.Left = margin
    fraHost.Top = lblTitle.Top + lblTitle.Height + 6
    fraHost.Width = Me.InsideWidth - margin * 2
    fraHost.Height = buttonTop - fraHost.Top - 6

    cmdCancel.Top = buttonTop
    cmdSave.Top = buttonTop

    cmdCancel.Left = Me.InsideWidth - cmdCancel.Width - margin
    cmdSave.Left = cmdCancel.Left - cmdSave.Width - 6
End Sub

Private Sub ClearDynamicControls()
    Dim i As Long

    For i = fraHost.Controls.Count - 1 To 0 Step -1
        fraHost.Controls.Remove fraHost.Controls(i).Name
    Next i

    Set mControlMap = CreateObject("Scripting.Dictionary")
End Sub

Private Sub BuildControls()
    Const SHEET_FIELDS As String = "xe.fields"
    Const ROW_HEIGHT As Single = 18
    Const ROW_GAP As Single = 6

    Dim wsFields As Worksheet

    Dim colFormID As Long
    Dim colDisplayOrder As Long
    Dim colFieldName As Long
    Dim colLabel As Long
    Dim colControlType As Long
    Dim colDataType As Long
    Dim colListID As Long

    Dim lastRow As Long
    Dim iRow As Long

    Dim sFormID As String
    Dim sFieldName As String
    Dim sLabel As String
    Dim sControlType As String
    Dim sDataType As String
    Dim sControlName As String
    Dim sListID As String

    Dim lbl As MSForms.Label
    Dim edt As MSForms.control

    Dim widthFrame As Single
    Dim margin As Single
    Dim curTop As Single
    Dim labelWidth As Single
    Dim inputWidth As Single

    If Not WorksheetExists(SHEET_FIELDS) Then
        MsgBox SHEET_FIELDS & " does not exist. Go speak to Mike.", vbExclamation, "xlEventing"
        Exit Sub
    End If

    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)

    colFormID = FindColumnInSheet(wsFields, "FormID")
    colDisplayOrder = FindColumnInSheet(wsFields, "DisplayOrder")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    colLabel = FindColumnInSheet(wsFields, "Label")
    colControlType = FindColumnInSheet(wsFields, "ControlType")
    colDataType = FindColumnInSheet(wsFields, "DataType")
    colListID = FindColumnInSheet(wsFields, "ListID")
    
    If (colFormID = 0) Or _
       (colDisplayOrder = 0) Or _
       (colFieldName = 0) Or _
       (colLabel = 0) Or _
       (colControlType = 0) Or _
       (colDataType = 0) Or _
       (colListID = 0) Then

        MsgBox "There are missing columns on " & SHEET_FIELDS & ". Go speak to Mike.", vbExclamation, "xlEventing"
        Exit Sub
    End If

    If mActiveRow >= 2 Then
        Me.Caption = "xlEventing: Edit"
    Else
        Me.Caption = "xlEventing: Append"
    End If

    lblTitle.Caption = mFormID

    margin = 12
    curTop = margin

    widthFrame = fraHost.InsideWidth
    If widthFrame <= 0 Then widthFrame = fraHost.Width

    labelWidth = widthFrame * 0.5
    inputWidth = widthFrame * 0.5

    lastRow = LastUsedRow(wsFields)
    If lastRow < 2 Then
        fraHost.ScrollBars = fmScrollBarsVertical
        fraHost.ScrollHeight = curTop + margin
        Exit Sub
    End If

    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsFields.Cells(iRow, colFormID).Value))

        If StrComp(sFormID, mFormID, vbTextCompare) = 0 Then
            sFieldName = Trim$(CStr(wsFields.Cells(iRow, colFieldName).Value))
            sLabel = Trim$(CStr(wsFields.Cells(iRow, colLabel).Value))
            sControlType = LCase$(Trim$(CStr(wsFields.Cells(iRow, colControlType).Value)))
            sDataType = LCase$(Trim$(CStr(wsFields.Cells(iRow, colDataType).Value)))
            sListID = Trim$(CStr(wsFields.Cells(iRow, colListID).Value))
            
            If Len(sFieldName) = 0 Then
                ' Skip blank field definitions
            Else
                sControlName = SafeControlSuffix(sFieldName, iRow)

                Set lbl = fraHost.Controls.Add("Forms.Label.1", "lbl_" & sControlName, True)
                Set edt = AddInputControl(sControlType, sControlName, sFieldName)
                
                lbl.Caption = sLabel
                lbl.Tag = sFieldName
                lbl.Left = margin
                lbl.Top = curTop + 2
                lbl.Width = labelWidth - (1.5 * margin)
                lbl.TextAlign = fmTextAlignRight
                
                edt.Tag = sFieldName & "|" & sListID
                edt.Left = lbl.Left + lbl.Width + (0.5 * margin)
                
                If sControlType = "checkbox" Then
                    edt.Top = curTop - 2
                    edt.Width = 18
                Else
                    edt.Top = curTop
                    edt.Width = inputWidth - (2 * margin)
                End If
                
                ConfigureInputControl edt, sControlType, sDataType

                If Not mControlMap.Exists(sFieldName) Then
                    mControlMap.Add sFieldName, edt.Name
                End If

                curTop = curTop + ROW_HEIGHT + ROW_GAP
            End If
        End If
    Next iRow

    fraHost.ScrollBars = fmScrollBarsVertical
    fraHost.KeepScrollBarsVisible = fmScrollBarsVertical
    fraHost.ScrollHeight = curTop + margin
End Sub

Private Function AddInputControl(ByVal pControlType As String, ByVal pControlSuffix As String, ByVal pFieldName As String) As MSForms.control
    Dim sControlType As String
    
    sControlType = LCase$(Trim$(pControlType))
    
    Select Case sControlType
        Case "textbox"
            Set AddInputControl = fraHost.Controls.Add("Forms.TextBox.1", "txt_" & pControlSuffix, True)

        Case "combo"
            Set AddInputControl = fraHost.Controls.Add("Forms.ComboBox.1", "cbo_" & pControlSuffix, True)
            
            Dim handler As clsComboHandler
            Set handler = New clsComboHandler
            handler.Init AddInputControl, pFieldName, Me
            
            mEventHandlers.Add handler

        Case "checkbox"
            Set AddInputControl = fraHost.Controls.Add("Forms.CheckBox.1", "chk_" & pControlSuffix, True)

        Case Else
            Set AddInputControl = fraHost.Controls.Add("Forms.TextBox.1", "txt_" & pControlSuffix, True)
    End Select
End Function

Private Sub PopulateAllLists()
    Const SHEET_FIELDS As String = "xe.fields"
    
    Dim wsFields As Worksheet
    Dim colFormID As Long
    Dim colFieldName As Long
    Dim colControlType As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    Dim sFieldName As String
    Dim sControlType As String
    
    If Not WorksheetExists(SHEET_FIELDS) Then Exit Sub
    
    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)
    
    colFormID = FindColumnInSheet(wsFields, "FormID")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    colControlType = FindColumnInSheet(wsFields, "ControlType")
    
    If (colFormID = 0) Or (colFieldName = 0) Or (colControlType = 0) Then Exit Sub
    
    lastRow = LastUsedRow(wsFields)
    If lastRow < 2 Then Exit Sub
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsFields.Cells(iRow, colFormID).Value))
        
        If StrComp(sFormID, mFormID, vbTextCompare) = 0 Then
            sFieldName = Trim$(CStr(wsFields.Cells(iRow, colFieldName).Value))
            sControlType = LCase$(Trim$(CStr(wsFields.Cells(iRow, colControlType).Value)))
            
            If (sControlType = "combo") And (Len(sFieldName) > 0) Then
                PopulateListForField sFieldName, True
            End If
        End If
    Next iRow
End Sub

Private Sub PopulateListForField(ByVal pFieldName As String, ByVal pPreserveValue As Boolean)
    Const SHEET_FIELDS As String = "xe.fields"
    
    Dim wsFields As Worksheet
    Dim colFormID As Long
    Dim colFieldName As Long
    Dim colControlType As Long
    Dim colListID As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    Dim sFieldName As String
    Dim sControlType As String
    Dim sListID As String
    
    Dim ctlName As String
    Dim ctl As MSForms.control
    Dim cbo As MSForms.ComboBox
    
    Dim values As Collection
    Dim vItem As Variant
    Dim oldValue As String
    
    If Not WorksheetExists(SHEET_FIELDS) Then Exit Sub
    If mControlMap Is Nothing Then Exit Sub
    If Not mControlMap.Exists(pFieldName) Then Exit Sub
    
    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)
    
    colFormID = FindColumnInSheet(wsFields, "FormID")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    colControlType = FindColumnInSheet(wsFields, "ControlType")
    colListID = FindColumnInSheet(wsFields, "ListID")
    
    If (colFormID = 0) Or (colFieldName = 0) Or (colControlType = 0) Or (colListID = 0) Then Exit Sub
    
    sListID = ""
    lastRow = LastUsedRow(wsFields)
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsFields.Cells(iRow, colFormID).Value))
        sFieldName = Trim$(CStr(wsFields.Cells(iRow, colFieldName).Value))
        sControlType = LCase$(Trim$(CStr(wsFields.Cells(iRow, colControlType).Value)))
        
        If StrComp(sFormID, mFormID, vbTextCompare) = 0 Then
            If StrComp(sFieldName, pFieldName, vbTextCompare) = 0 Then
                If sControlType = "combo" Then
                    sListID = Trim$(CStr(wsFields.Cells(iRow, colListID).Value))
                End If
                Exit For
            End If
        End If
    Next iRow
    
    If Len(sListID) = 0 Then Exit Sub
    
    ctlName = CStr(mControlMap(pFieldName))
    Set ctl = fraHost.Controls(ctlName)
    
    If TypeName(ctl) <> "ComboBox" Then Exit Sub
    Set cbo = ctl
    
    oldValue = Trim$(CStr(cbo.Value))
    
    cbo.Clear
    
    Set values = GetListValues(sListID)
    If Not values Is Nothing Then
        For Each vItem In values
            cbo.AddItem CStr(vItem)
        Next vItem
    End If
    
    If pPreserveValue Then
        If ComboContains(cbo, oldValue) Then
            cbo.Value = oldValue
        Else
            cbo.ListIndex = -1
        End If
    Else
        If ComboContains(cbo, oldValue) Then
            cbo.Value = oldValue
        Else
            cbo.ListIndex = -1
        End If
    End If
End Sub

Public Sub RefreshChildLists(ByVal pParentFieldName As String)
    Const SHEET_FIELDS As String = "xe.fields"
    
    Dim wsFields As Worksheet
    Dim colFormID As Long
    Dim colFieldName As Long
    Dim colControlType As Long
    Dim colParentField1 As Long
    Dim colParentField2 As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    Dim sFieldName As String
    Dim sControlType As String
    Dim sParentField1 As String
    Dim sParentField2 As String
    
    If Not WorksheetExists(SHEET_FIELDS) Then Exit Sub
    
    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)
    
    colFormID = FindColumnInSheet(wsFields, "FormID")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    colControlType = FindColumnInSheet(wsFields, "ControlType")
    colParentField1 = FindColumnInSheet(wsFields, "ParentField1")
    colParentField2 = FindColumnInSheet(wsFields, "ParentField2")
    
    If (colFormID = 0) Or (colFieldName = 0) Or (colControlType = 0) Then Exit Sub
    
    lastRow = LastUsedRow(wsFields)
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsFields.Cells(iRow, colFormID).Value))
        
        If StrComp(sFormID, mFormID, vbTextCompare) = 0 Then
            sFieldName = Trim$(CStr(wsFields.Cells(iRow, colFieldName).Value))
            sControlType = LCase$(Trim$(CStr(wsFields.Cells(iRow, colControlType).Value)))
            
            If colParentField1 > 0 Then sParentField1 = Trim$(CStr(wsFields.Cells(iRow, colParentField1).Value)) Else sParentField1 = ""
            If colParentField2 > 0 Then sParentField2 = Trim$(CStr(wsFields.Cells(iRow, colParentField2).Value)) Else sParentField2 = ""
            
            If sControlType = "combo" Then
                If (StrComp(sParentField1, pParentFieldName, vbTextCompare) = 0) Or _
                   (StrComp(sParentField2, pParentFieldName, vbTextCompare) = 0) Then
                    PopulateListForField sFieldName, False
                End If
            End If
        End If
    Next iRow
End Sub

Private Sub LoadRowValues(ByVal pLoadDependentFields As Boolean)
    Const SHEET_FIELDS As String = "xe.fields"
    
    Dim wsTarget As Worksheet
    Dim wsFields As Worksheet
    
    Dim sTargetSheet As String
    Dim vFieldName As Variant
    Dim sControlName As String
    Dim ctl As MSForms.control
    Dim lCol As Long
    Dim vValue As Variant
    
    Dim colFormID As Long
    Dim colFieldName As Long
    Dim colParentField1 As Long
    Dim colParentField2 As Long
    Dim lastRowFields As Long
    Dim iRow As Long
    
    Dim sFormID As String
    Dim sFieldName As String
    Dim sParentField1 As String
    Dim sParentField2 As String
    Dim hasParents As Boolean
    
    If mActiveRow < 2 Then Exit Sub
    If mControlMap Is Nothing Then Exit Sub
    
    sTargetSheet = GetTargetSheetForForm(mFormID)
    
    If Len(sTargetSheet) = 0 Then
        MsgBox "No TargetSheet defined for form '" & mFormID & "'.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    If Not WorksheetExists(sTargetSheet) Then
        MsgBox "Target sheet '" & sTargetSheet & "' does not exist.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    If Not WorksheetExists(SHEET_FIELDS) Then
        MsgBox "xe.fields does not exist.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Set wsTarget = ActiveWorkbook.Worksheets(sTargetSheet)
    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)
    
    colFormID = FindColumnInSheet(wsFields, "FormID")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    colParentField1 = FindColumnInSheet(wsFields, "ParentField1")
    colParentField2 = FindColumnInSheet(wsFields, "ParentField2")
    
    If (colFormID = 0) Or (colFieldName = 0) Then Exit Sub
    
    lastRowFields = LastUsedRow(wsFields)
    
    For iRow = 2 To lastRowFields
        sFormID = Trim$(CStr(wsFields.Cells(iRow, colFormID).Value))
        sFieldName = Trim$(CStr(wsFields.Cells(iRow, colFieldName).Value))
        
        If StrComp(sFormID, mFormID, vbTextCompare) = 0 Then
            sParentField1 = ""
            sParentField2 = ""
            
            If colParentField1 > 0 Then sParentField1 = Trim$(CStr(wsFields.Cells(iRow, colParentField1).Value))
            If colParentField2 > 0 Then sParentField2 = Trim$(CStr(wsFields.Cells(iRow, colParentField2).Value))
            
            hasParents = (Len(sParentField1) > 0) Or (Len(sParentField2) > 0)
            
            If hasParents = pLoadDependentFields Then
                If mControlMap.Exists(sFieldName) Then
                    sControlName = CStr(mControlMap(sFieldName))
                    Set ctl = fraHost.Controls(sControlName)
                    
                    lCol = FindColumnInSheet(wsTarget, sFieldName)
                    If lCol > 0 Then
                        vValue = wsTarget.Cells(mActiveRow, lCol).Value
                        SetControlValue ctl, vValue
                    End If
                End If
            End If
        End If
    Next iRow
End Sub

Private Function GetListValues(ByVal pListID As String) As Collection
    Const SHEET_LISTS As String = "xe.lists"
    
    Dim wsLists As Worksheet
    Dim wsSource As Worksheet
    
    Dim colListID As Long
    Dim colSourceSheet As Long
    Dim colValueField As Long
    Dim colFilterField1 As Long
    Dim colFilterParentField1 As Long
    Dim colFilterField2 As Long
    Dim colFilterParentField2 As Long
    Dim colFilterField3 As Long
    Dim colFilterParentField3 As Long
    Dim colDistinctValues As Long
    Dim colSortValues As Long
    
    Dim lastRowLists As Long
    Dim iRow As Long
    
    Dim sListID As String
    Dim sSourceSheet As String
    Dim sValueField As String
    Dim sFilterField1 As String
    Dim sFilterParentField1 As String
    Dim sFilterField2 As String
    Dim sFilterParentField2 As String
    Dim sFilterField3 As String
    Dim sFilterParentField3 As String
    Dim sDistinctValues As String
    Dim sSortValues As String
    
    Dim srcColValue As Long
    Dim srcColFilter1 As Long
    Dim srcColFilter2 As Long
    Dim srcColFilter3 As Long
    
    Dim lastRowSource As Long
    Dim srcRow As Long
    
    Dim parentValue1 As String
    Dim parentValue2 As String
    Dim parentValue3 As String
    
    Dim currentValue As String
    Dim includeRow As Boolean
    
    Dim outValues As Collection
    Dim dictDistinct As Object
    
    If Not WorksheetExists(SHEET_LISTS) Then
        Set GetListValues = Nothing
        Exit Function
    End If
    
    Set wsLists = ActiveWorkbook.Worksheets(SHEET_LISTS)
    
    colListID = FindColumnInSheet(wsLists, "ListID")
    colSourceSheet = FindColumnInSheet(wsLists, "SourceSheet")
    colValueField = FindColumnInSheet(wsLists, "ValueField")
    colFilterField1 = FindColumnInSheet(wsLists, "FilterField1")
    colFilterParentField1 = FindColumnInSheet(wsLists, "FilterParentField1")
    colFilterField2 = FindColumnInSheet(wsLists, "FilterField2")
    colFilterParentField2 = FindColumnInSheet(wsLists, "FilterParentField2")
    colFilterField3 = FindColumnInSheet(wsLists, "FilterField3")
    colFilterParentField3 = FindColumnInSheet(wsLists, "FilterParentField3")
    colDistinctValues = FindColumnInSheet(wsLists, "DistinctValues")
    colSortValues = FindColumnInSheet(wsLists, "SortValues")
    
    If (colListID = 0) Or (colSourceSheet = 0) Or (colValueField = 0) Then
        Set GetListValues = Nothing
        Exit Function
    End If
    
    lastRowLists = LastUsedRow(wsLists)
    If lastRowLists < 2 Then
        Set GetListValues = Nothing
        Exit Function
    End If
    
    sSourceSheet = ""
    
    For iRow = 2 To lastRowLists
        sListID = Trim$(CStr(wsLists.Cells(iRow, colListID).Value))
        
        If StrComp(sListID, pListID, vbTextCompare) = 0 Then
            sSourceSheet = Trim$(CStr(wsLists.Cells(iRow, colSourceSheet).Value))
            sValueField = Trim$(CStr(wsLists.Cells(iRow, colValueField).Value))
            
            If colFilterField1 > 0 Then sFilterField1 = Trim$(CStr(wsLists.Cells(iRow, colFilterField1).Value))
            If colFilterParentField1 > 0 Then sFilterParentField1 = Trim$(CStr(wsLists.Cells(iRow, colFilterParentField1).Value))
            If colFilterField2 > 0 Then sFilterField2 = Trim$(CStr(wsLists.Cells(iRow, colFilterField2).Value))
            If colFilterParentField2 > 0 Then sFilterParentField2 = Trim$(CStr(wsLists.Cells(iRow, colFilterParentField2).Value))
            If colFilterField3 > 0 Then sFilterField3 = Trim$(CStr(wsLists.Cells(iRow, colFilterField3).Value))
            If colFilterParentField3 > 0 Then sFilterParentField3 = Trim$(CStr(wsLists.Cells(iRow, colFilterParentField3).Value))
            If colDistinctValues > 0 Then sDistinctValues = Trim$(CStr(wsLists.Cells(iRow, colDistinctValues).Value))
            If colSortValues > 0 Then sSortValues = Trim$(CStr(wsLists.Cells(iRow, colSortValues).Value))
            
            Exit For
        End If
    Next iRow
    
    If Len(sSourceSheet) = 0 Then
        Set GetListValues = Nothing
        Exit Function
    End If
    
    If Not WorksheetExists(sSourceSheet) Then
        Set GetListValues = Nothing
        Exit Function
    End If
    
    Set wsSource = ActiveWorkbook.Worksheets(sSourceSheet)
    
    srcColValue = FindColumnInSheet(wsSource, sValueField)
    If srcColValue = 0 Then
        Set GetListValues = Nothing
        Exit Function
    End If
    
    If Len(sFilterField1) > 0 Then
        srcColFilter1 = FindColumnInSheet(wsSource, sFilterField1)
        If srcColFilter1 = 0 Then
            Set GetListValues = Nothing
            Exit Function
        End If
    End If
    
    If Len(sFilterField2) > 0 Then
        srcColFilter2 = FindColumnInSheet(wsSource, sFilterField2)
        If srcColFilter2 = 0 Then
            Set GetListValues = Nothing
            Exit Function
        End If
    End If
    
    If Len(sFilterField3) > 0 Then
        srcColFilter3 = FindColumnInSheet(wsSource, sFilterField3)
        If srcColFilter3 = 0 Then
            Set GetListValues = Nothing
            Exit Function
        End If
    End If
    
    If Len(sFilterField1) > 0 Then
        If Len(sFilterParentField1) = 0 Then sFilterParentField1 = sFilterField1
        parentValue1 = GetFormFieldValue(sFilterParentField1)
    End If
    
    If Len(sFilterField2) > 0 Then
        If Len(sFilterParentField2) = 0 Then sFilterParentField2 = sFilterField2
        parentValue2 = GetFormFieldValue(sFilterParentField2)
    End If
    
    If Len(sFilterField3) > 0 Then
        If Len(sFilterParentField3) = 0 Then sFilterParentField3 = sFilterField3
        parentValue3 = GetFormFieldValue(sFilterParentField3)
    End If
    
    Set outValues = New Collection
    
    If IsYesValue(sDistinctValues) Then
        Set dictDistinct = CreateObject("Scripting.Dictionary")
    End If
    
    lastRowSource = LastUsedRow(wsSource)
    
    For srcRow = 2 To lastRowSource
        includeRow = True
        
        If Len(sFilterField1) > 0 Then
            If StrComp(Trim$(CStr(wsSource.Cells(srcRow, srcColFilter1).Value)), parentValue1, vbTextCompare) <> 0 Then
                includeRow = False
            End If
        End If
        
        If includeRow And Len(sFilterField2) > 0 Then
            If StrComp(Trim$(CStr(wsSource.Cells(srcRow, srcColFilter2).Value)), parentValue2, vbTextCompare) <> 0 Then
                includeRow = False
            End If
        End If
        
        If includeRow And Len(sFilterField3) > 0 Then
            If StrComp(Trim$(CStr(wsSource.Cells(srcRow, srcColFilter3).Value)), parentValue3, vbTextCompare) <> 0 Then
                includeRow = False
            End If
        End If
        
        If includeRow Then
            currentValue = Trim$(CStr(wsSource.Cells(srcRow, srcColValue).Value))
            
            If Len(currentValue) > 0 Then
                If dictDistinct Is Nothing Then
                    outValues.Add currentValue
                Else
                    If Not dictDistinct.Exists(LCase$(currentValue)) Then
                        dictDistinct.Add LCase$(currentValue), currentValue
                        outValues.Add currentValue
                    End If
                End If
            End If
        End If
    Next srcRow
    
    If IsYesValue(sSortValues) Then
        Set outValues = SortCollectionText(outValues)
    End If
    
    Set GetListValues = outValues
End Function

Private Function GetFormFieldValue(ByVal pFieldName As String) As String
    Dim ctlName As String
    Dim ctl As MSForms.control
    
    GetFormFieldValue = ""
    
    If mControlMap Is Nothing Then Exit Function
    If Not mControlMap.Exists(pFieldName) Then Exit Function
    
    ctlName = CStr(mControlMap(pFieldName))
    Set ctl = fraHost.Controls(ctlName)
    
    Select Case TypeName(ctl)
        Case "TextBox"
            GetFormFieldValue = Trim$(CStr(ctl.Text))
        Case "ComboBox"
            GetFormFieldValue = Trim$(CStr(ctl.Value))
        Case "CheckBox"
            GetFormFieldValue = Trim$(CStr(ctl.Value))
    End Select
End Function

Private Function IsYesValue(ByVal pValue As String) As Boolean
    Dim s As String
    
    s = LCase$(Trim$(pValue))
    
    IsYesValue = (s = "y") Or (s = "yes") Or (s = "true") Or (s = "1")
End Function

Private Function SortCollectionText(ByVal pValues As Collection) As Collection
    Dim arr() As String
    Dim i As Long
    Dim j As Long
    Dim temp As String
    Dim outValues As Collection
    
    Set outValues = New Collection
    
    If pValues Is Nothing Then
        Set SortCollectionText = outValues
        Exit Function
    End If
    
    If pValues.Count = 0 Then
        Set SortCollectionText = outValues
        Exit Function
    End If
    
    ReDim arr(1 To pValues.Count)
    
    For i = 1 To pValues.Count
        arr(i) = CStr(pValues(i))
    Next i
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    For i = LBound(arr) To UBound(arr)
        outValues.Add arr(i)
    Next i
    
    Set SortCollectionText = outValues
End Function

Private Sub ConfigureInputControl(ByVal pCtl As MSForms.control, ByVal pControlType As String, ByVal pDataType As String)

    If Not IsSupportedControlType(pControlType) Then
        MsgBox "Control Type " & pControlType & " is not yet supported", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Select Case LCase$(Trim$(pControlType))
        Case "textbox"
            Dim txt As MSForms.TextBox
            Set txt = pCtl

            txt.SpecialEffect = fmSpecialEffectSunken
            txt.TextAlign = fmTextAlignLeft

            Select Case pDataType
                Case "int", "integer", "long"
                    ' Placeholder for later numeric validation
                Case "float", "double", "single", "number", "numeric", "percent", "percentage"
                    ' Placeholder for later numeric validation
                Case "date"
                    ' Placeholder for later date validation
            End Select

        Case "combo"
            Dim cbo As MSForms.ComboBox
            Set cbo = pCtl

            cbo.Style = fmStyleDropDownCombo
            cbo.MatchEntry = fmMatchEntryComplete
            cbo.ListRows = 12
            
        
        Case "checkbox"
            Dim chk As MSForms.CheckBox
            Set chk = pCtl
            
            chk.Caption = ""
            chk.Value = False
            chk.BackStyle = fmBackStyleTransparent
    End Select
End Sub

Private Sub SaveFormToRow()
    Dim wsTarget As Worksheet
    Dim sTargetSheet As String
    Dim lTargetRow As Long
    
    On Error GoTo ErrHandler
    
    sTargetSheet = GetTargetSheetForForm(mFormID)
    
    If Len(sTargetSheet) = 0 Then
        MsgBox "No TargetSheet defined for form '" & mFormID & "' in xe.forms.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Set wsTarget = EnsureTargetSheetExists(sTargetSheet, mFormID)
    If wsTarget Is Nothing Then Exit Sub
    
    If mActiveRow < 0 Then
        lTargetRow = LastUsedRow(wsTarget) + 1
        If lTargetRow < 2 Then lTargetRow = 2
    Else
        If mActiveRow < 2 Then
            MsgBox "Invalid target row " & CStr(mActiveRow) & ".", vbExclamation, "xlEventing"
            Exit Sub
        End If
        lTargetRow = mActiveRow
    End If
    
    WriteFormValuesToSheet wsTarget, lTargetRow
    
    wsTarget.Activate
    wsTarget.Cells(lTargetRow, 1).Select
    
    Unload Me
    Exit Sub

ErrHandler:
    MsgBox "Error saving form: " & Err.Description, vbExclamation, "xlEventing"
End Sub

Private Sub WriteFormValuesToSheet(ByVal pWS As Worksheet, ByVal pTargetRow As Long)
    Dim vFieldName As Variant
    Dim sControlName As String
    Dim ctl As MSForms.control
    Dim lCol As Long
    Dim vValue As Variant
    
    If mControlMap Is Nothing Then Exit Sub
    
    For Each vFieldName In mControlMap.Keys
        sControlName = CStr(mControlMap(vFieldName))
        Set ctl = fraHost.Controls(sControlName)
        
        lCol = FindColumnInSheet(pWS, CStr(vFieldName))
        
        If lCol = 0 Then
            MsgBox "Target sheet '" & pWS.Name & "' is missing column '" & CStr(vFieldName) & "'.", vbExclamation, "xlEventing"
            Exit Sub
        End If
        
        vValue = GetControlValue(ctl)
        pWS.Cells(pTargetRow, lCol).Value = vValue
    Next vFieldName
End Sub
