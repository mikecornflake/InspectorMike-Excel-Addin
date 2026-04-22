Attribute VB_Name = "appXLEventing"

Option Explicit

Public Sub ShowXlEventingForm(ByVal pFormID As String, ByVal pActiveRow As Long)
    Unload frmXLEventing
    Load frmXLEventing
    frmXLEventing.SetupForm pFormID, pActiveRow
    frmXLEventing.Show
End Sub

Public Sub ShowXLAdminForm()
    Unload frmXLAdmin
    Load frmXLAdmin
    frmXLAdmin.Show
End Sub

Public Sub ShowXlLaunchForm()
    Unload frmXLLaunch
    Load frmXLLaunch
    frmXLLaunch.Show
End Sub

Public Sub ShowXlEventingForm_EditOrAppendFromActiveSheet()
    Dim ws As Worksheet
    Dim sFormID As String
    Dim sFormType As String
    Dim lRow As Long
    
    Set ws = ActiveSheet
    If ws Is Nothing Then
        ShowXlLaunchForm
        Exit Sub
    End If
    
    sFormID = GetFormIDForTargetSheet(ws.Name)
    
    If Len(sFormID) = 0 Then
        ShowXlLaunchForm
        Exit Sub
    End If
    
    sFormType = LCase$(Trim$(GetFormTypeForForm(sFormID)))
    
    Select Case sFormType
        Case "configuration"
            MsgBox "Please edit the Excel sheet directly", vbInformation, "xlEventing"
            Exit Sub
            
        Case "event"
            lRow = ActiveCell.Row
            
            If lRow < 2 Then
                ShowXlEventingForm sFormID, -1
                Exit Sub
            End If
            
            If IsWorksheetRowPopulated(ws, lRow) Then
                ShowXlEventingForm sFormID, lRow
            Else
                ShowXlEventingForm sFormID, -1
            End If
            
        Case Else
            ShowXlLaunchForm
    End Select
End Sub

Public Function GetFormTypeForForm(ByVal pFormID As String) As String
    Const SHEET_FORMS As String = "xe.forms"
    
    Dim wsForms As Worksheet
    Dim colFormID As Long
    Dim colType As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    
    GetFormTypeForForm = ""
    
    If Not WorksheetExists(SHEET_FORMS) Then Exit Function
    
    Set wsForms = ActiveWorkbook.Worksheets(SHEET_FORMS)
    
    colFormID = FindColumnInSheet(wsForms, "FormID")
    colType = FindColumnInSheet(wsForms, "Type")
    
    If (colFormID = 0) Or (colType = 0) Then Exit Function
    
    lastRow = LastUsedRow(wsForms)
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsForms.Cells(iRow, colFormID).Value))
        
        If StrComp(sFormID, pFormID, vbTextCompare) = 0 Then
            GetFormTypeForForm = Trim$(CStr(wsForms.Cells(iRow, colType).Value))
            Exit Function
        End If
    Next iRow
End Function

' TODO Move this to LibraryWorksheet
Public Function IsWorksheetRowPopulated(ByVal pWS As Worksheet, ByVal pRow As Long) As Boolean
    Dim lastCol As Long
    Dim iCol As Long
    
    IsWorksheetRowPopulated = False
    
    If pRow < 2 Then Exit Function
    
    lastCol = pWS.Cells(1, pWS.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function
    
    For iCol = 1 To lastCol
        If Len(Trim$(CStr(pWS.Cells(pRow, iCol).Value))) > 0 Then
            IsWorksheetRowPopulated = True
            Exit Function
        End If
    Next iCol
End Function

Public Function SafeControlSuffix(ByVal pFieldName As String, ByVal pRow As Long) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    s = Trim$(pFieldName)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch Like "[A-Za-z0-9_]" Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next i

    If Len(out) = 0 Then
        out = "Field_" & CStr(pRow)
    End If

    SafeControlSuffix = out & "_" & CStr(pRow)
End Function

' TODO: Move this to LibraryWorksheets
Public Function WorksheetExists(ByVal pSheetName As String) As Boolean
    Dim ws As Worksheet

    WorksheetExists = False

    For Each ws In ActiveWorkbook.Worksheets
        If StrComp(ws.Name, pSheetName, vbTextCompare) = 0 Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function

' TODO: Move this to LibraryWorksheets
Public Function FindColumnInSheet(ByVal pWS As Worksheet, ByVal pHeaderName As String) As Long
    Dim lastCol As Long
    Dim iCol As Long
    Dim sHeader As String

    FindColumnInSheet = 0

    lastCol = pWS.Cells(1, pWS.Columns.Count).End(xlToLeft).Column

    For iCol = 1 To lastCol
        sHeader = Trim$(CStr(pWS.Cells(1, iCol).Value))
        If StrComp(sHeader, pHeaderName, vbTextCompare) = 0 Then
            FindColumnInSheet = iCol
            Exit Function
        End If
    Next iCol
End Function

' TODO: Move this to LibraryWorksheets
Public Function LastUsedRow(ByVal pWS As Worksheet) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = pWS.Cells.Find(What:="*", _
                                  After:=pWS.Cells(1, 1), _
                                  LookIn:=xlFormulas, _
                                  LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = lastCell.Row
    End If
End Function

Public Sub BasicTidyWorksheet(ByVal pWS As Worksheet)
    pWS.Activate
    BasicTidy
End Sub


' Thanks Copilot 8/4/2026
Sub IntelligentlyInsertDateTime()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim activeRow As Long
    activeRow = ActiveCell.Row
    
    ' Do nothing if user is on header row
    If activeRow = 1 Then Exit Sub
    
    Dim colDate As Long
    Dim colTimeLocal As Long
    Dim colStartTime As Long
    Dim colEndTime As Long
    
    colDate = FindFirstColumn(Array("Date"))
    colTimeLocal = FindFirstColumn(Array("Time (Local)", "Time"))
    colStartTime = FindFirstColumn(Array("Start Time (Local)", "Start Time"))
    colEndTime = FindFirstColumn(Array("End Time (Local)", "End Time"))
    
    ' --- Date ---
    If colDate > 0 Then
        If isEmpty(ws.Cells(activeRow, colDate)) Then
            ws.Cells(activeRow, colDate).Value = Date
        End If
    End If
    
    ' --- Time (Local) ---
    If colTimeLocal > 0 Then
        If isEmpty(ws.Cells(activeRow, colTimeLocal)) Then
            ws.Cells(activeRow, colTimeLocal).Value = Time
        End If
    End If
    
    ' --- Start / End Time Logic ---
    If colStartTime > 0 Then
        
        ' If Start Time exists and is blank ? set to now
        If isEmpty(ws.Cells(activeRow, colStartTime)) Then
            ws.Cells(activeRow, colStartTime).Value = Time
        
        ' If Start Time exists and NOT blank
        Else
            ' Then if End Time exists and is blank ? set to now
            If colEndTime > 0 Then
                If isEmpty(ws.Cells(activeRow, colEndTime)) Then
                    ws.Cells(activeRow, colEndTime).Value = Time
                End If
            End If
        End If
    End If
End Sub

Public Function GetTargetSheetForForm(ByVal pFormID As String) As String
    Const SHEET_FORMS As String = "xe.forms"
    
    Dim wsForms As Worksheet
    Dim colFormID As Long
    Dim colTargetSheet As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    
    GetTargetSheetForForm = ""
    
    If Not WorksheetExists(SHEET_FORMS) Then Exit Function
    
    Set wsForms = ActiveWorkbook.Worksheets(SHEET_FORMS)
    
    colFormID = FindColumnInSheet(wsForms, "FormID")
    colTargetSheet = FindColumnInSheet(wsForms, "TargetSheet")
    
    If (colFormID = 0) Or (colTargetSheet = 0) Then Exit Function
    
    lastRow = LastUsedRow(wsForms)
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsForms.Cells(iRow, colFormID).Value))
        
        If StrComp(sFormID, pFormID, vbTextCompare) = 0 Then
            GetTargetSheetForForm = Trim$(CStr(wsForms.Cells(iRow, colTargetSheet).Value))
            Exit Function
        End If
    Next iRow
End Function

Public Function EnsureTargetSheetExists(ByVal pSheetName As String, ByVal pFormID As String) As Worksheet
    Dim ws As Worksheet
    
    If WorksheetExists(pSheetName) Then
        Set ws = ActiveWorkbook.Worksheets(pSheetName)
        
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
    Else
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = pSheetName
        ws.Move After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        ColourEventTab ws
        
        CreateSheetHeadersFromFields ws, pFormID
        
        On Error Resume Next
        BasicTidyWorksheet ws
        On Error GoTo 0
    End If
    
    Set EnsureTargetSheetExists = ws
End Function

Public Sub ColourAdminTab(ByVal pSheet As Worksheet)
    With pSheet.Tab
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
End Sub

Public Sub ColourEventTab(ByVal pSheet As Worksheet)
    With pSheet.Tab
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        
'        .ColorIndex = xlColorIndexNone
'        .TintAndShade = 0
    End With
End Sub

Public Sub CreateSheetHeadersFromFields(ByVal pWS As Worksheet, ByVal pFormID As String)
    Const SHEET_FIELDS As String = "xe.fields"
    
    Dim wsFields As Worksheet
    Dim colFormID As Long
    Dim colDisplayOrder As Long
    Dim colFieldName As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    Dim sFieldName As String
    
    Dim colHeaders As Collection
    Dim vItem As Variant
    Dim iCol As Long
    
    If Not WorksheetExists(SHEET_FIELDS) Then
        MsgBox "xe.fields not found.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Set wsFields = ActiveWorkbook.Worksheets(SHEET_FIELDS)
    
    colFormID = FindColumnInSheet(wsFields, "FormID")
    colDisplayOrder = FindColumnInSheet(wsFields, "DisplayOrder")
    colFieldName = FindColumnInSheet(wsFields, "FieldName")
    
    If (colFormID = 0) Or (colDisplayOrder = 0) Or (colFieldName = 0) Then
        MsgBox "xe.fields is missing required columns.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Set colHeaders = New Collection
    
    lastRow = LastUsedRow(wsFields)
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsFields.Cells(iRow, colFormID).Value))
        
        If StrComp(sFormID, pFormID, vbTextCompare) = 0 Then
            sFieldName = Trim$(CStr(wsFields.Cells(iRow, colFieldName).Value))
            
            If Len(sFieldName) > 0 Then
                colHeaders.Add Array( _
                    CLng(wsFields.Cells(iRow, colDisplayOrder).Value), _
                    sFieldName _
                )
            End If
        End If
    Next iRow
    
    Set colHeaders = CollectionBubbleSort(colHeaders, True)
    
    iCol = 1
    For Each vItem In colHeaders
        pWS.Cells(1, iCol).Value = vItem(1)
        iCol = iCol + 1
    Next vItem
    
    pWS.Rows(1).Font.Bold = True
    
    On Error Resume Next
    
    BasicTidyWorksheet pWS
    
    On Error GoTo 0
End Sub

' TODO Move this to a LibraryControls
Public Sub RemoveComboItem(ByVal cbo As MSForms.ComboBox, ByVal pValue As String)
    Dim i As Long
    
    For i = cbo.ListCount - 1 To 0 Step -1
        If StrComp(Trim$(cbo.List(i)), Trim$(pValue), vbTextCompare) = 0 Then
            cbo.RemoveItem i
            Exit For
        End If
    Next i
End Sub

' TODO put this in LibraryControls
Public Function ComboContains(ByVal cbo As MSForms.ComboBox, ByVal pValue As String) As Boolean
    Dim i As Long
    
    ComboContains = False
    
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), pValue, vbTextCompare) = 0 Then
            ComboContains = True
            Exit Function
        End If
    Next i
End Function

' TODO put this in LibraryControls
Public Function GetControlValue(ByVal pCtl As MSForms.control) As Variant

    Select Case TypeName(pCtl)
        Case "TextBox"
            GetControlValue = Trim$(CStr(pCtl.Text))
        
        Case "ComboBox"
            GetControlValue = Trim$(CStr(pCtl.Value))
        
        Case "CheckBox"
            GetControlValue = CBool(pCtl.Value)
        
        Case Else
            GetControlValue = Trim$(CStr(pCtl.Value))
    End Select
End Function

' TODO put this in LibraryControls
Public Sub SetControlValue(ByVal pCtl As MSForms.control, ByVal pValue As Variant)
    Select Case TypeName(pCtl)
        Case "TextBox"
            If IsNull(pValue) Then
                pCtl.Text = ""
            Else
                pCtl.Text = Trim$(CStr(pValue))
            End If
        
        Case "ComboBox"
            If IsNull(pValue) Then
                pCtl.Value = ""
            Else
                pCtl.Value = Trim$(CStr(pValue))
            End If
        
        Case "CheckBox"
            If IsNull(pValue) Or Len(Trim$(CStr(pValue))) = 0 Then
                pCtl.Value = False
            Else
                Select Case LCase$(Trim$(CStr(pValue)))
                    Case "true", "yes", "y", "1"
                        pCtl.Value = True
                    Case "false", "no", "n", "0"
                        pCtl.Value = False
                    Case Else
                        On Error Resume Next
                        pCtl.Value = CBool(pValue)
                        On Error GoTo 0
                End Select
            End If
    End Select
End Sub

' TODO put this in LibraryControls
Public Function IsSupportedControlType(ByVal pControlType As String) As Boolean
    Select Case LCase$(Trim$(pControlType))
        Case "textbox", "combo", "checkbox"
            IsSupportedControlType = True
        Case Else
            IsSupportedControlType = False
    End Select
End Function

Public Function GetFormIDForTargetSheet(ByVal pTargetSheet As String) As String
    Const SHEET_FORMS As String = "xe.forms"
    
    Dim wsForms As Worksheet
    Dim colFormID As Long
    Dim colTargetSheet As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    Dim sTargetSheet As String
    
    GetFormIDForTargetSheet = ""
    
    If Not WorksheetExists(SHEET_FORMS) Then Exit Function
    
    Set wsForms = ActiveWorkbook.Worksheets(SHEET_FORMS)
    
    colFormID = FindColumnInSheet(wsForms, "FormID")
    colTargetSheet = FindColumnInSheet(wsForms, "TargetSheet")
    
    If (colFormID = 0) Or (colTargetSheet = 0) Then Exit Function
    
    lastRow = LastUsedRow(wsForms)
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsForms.Cells(iRow, colFormID).Value))
        sTargetSheet = Trim$(CStr(wsForms.Cells(iRow, colTargetSheet).Value))
        
        If StrComp(sTargetSheet, pTargetSheet, vbTextCompare) = 0 Then
            GetFormIDForTargetSheet = sFormID
            Exit Function
        End If
    Next iRow
End Function
