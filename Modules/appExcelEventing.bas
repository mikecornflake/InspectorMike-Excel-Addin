Attribute VB_Name = "appExcelEventing"

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

Public Sub ShowXlEventingForm_Append(ByVal pFormID As String)
    ShowXlEventingForm pFormID, -1
End Sub

Public Sub ShowXlEventingForm_Edit(ByVal pFormID As String)
    ShowXlEventingForm pFormID, ActiveCell.Row
End Sub

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

