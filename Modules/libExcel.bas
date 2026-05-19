Attribute VB_Name = "libExcel"
' 24 Apr 2026 - refactored from libFiles
'            - Stopped using OneDrive/Sharepoint URL safe code.  MS broke the API too many times
' 18 May 2026 - Moved Workbook routines to libWorkbook
'             - Moved Worksheet routines to libWorksheet
'             - Rationalised namespace
'             - Added Selection_NextVisibleCellDown

Option Explicit
Option Private Module

' Opens the standard Save As dialog
Public Sub Application_Original_Save_As_Dialog()
    Application.Dialogs(xlDialogSaveAs).Show
End Sub

' Source - https://stackoverflow.com/a/78931790
' Posted by MasterPlexus
' Retrieved 2026-05-14, License - CC BY-SA 4.0
' Code review/Safety checks by CoPilot
Public Sub Selection_NextVisibleCellDown()
    Dim offsetRows As Long
    Dim nxt As Range

    offsetRows = 1
    Set nxt = ActiveCell.Offset(offsetRows, 0)

    ' Skip hidden rows
    Do While nxt.EntireRow.Hidden
        offsetRows = offsetRows + 1
        Set nxt = ActiveCell.Offset(offsetRows, 0)
        ' Optional: bail out if we run off the sheet
        If nxt.row > ActiveSheet.Rows.Count Then Exit Sub
    Loop

    nxt.Activate
End Sub

Public Sub Selection_TitleCase()
    Dim txtOnly As Range
    Dim cell As Range

    On Error Resume Next
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If txtOnly Is Nothing Then Exit Sub

    For Each cell In txtOnly.Cells
        cell.Value = StrConv(CStr(cell.Value), vbProperCase)
    Next cell
End Sub

Public Sub Selection_SentenceCase()
    Dim txtOnly As Range
    Dim cell As Range

    On Error Resume Next
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If txtOnly Is Nothing Then Exit Sub

    For Each cell In txtOnly.Cells
        cell.Value = Text_ToSentenceCase(CStr(cell.Value))
    Next cell
End Sub

Public Sub Selection_Uppercase()
    Dim txtOnly As Range
    Dim cell As Range

    On Error Resume Next
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If txtOnly Is Nothing Then Exit Sub

    For Each cell In txtOnly.Cells
        cell.Value = UCase$(CStr(cell.Value))
    Next cell
End Sub

Public Sub Selection_Lowercase()
    Dim txtOnly As Range
    Dim cell As Range

    On Error Resume Next
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0

    If txtOnly Is Nothing Then Exit Sub

    For Each cell In txtOnly.Cells
        cell.Value = LCase$(CStr(cell.Value))
    Next cell
End Sub
