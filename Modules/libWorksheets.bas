Attribute VB_Name = "libWorksheets"
Option Explicit
Option Private Module

Public Sub FreezeTopRow(ByVal pWS As Worksheet)
    pWS.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

Public Function AddSheet(ByVal pWB As Workbook, _
                         ByVal pName As String, _
                         Optional ByVal pIndex As Long = 1) As Worksheet
    Dim ws As Worksheet
    
    Set ws = FindSheet(pWB, pName)
    
    If ws Is Nothing Then
        Set ws = pWB.Worksheets.Add(After:=pWB.Worksheets(pWB.Worksheets.Count))
        ws.Name = pName
    End If
    
    Set AddSheet = ws
    
    If pWB.Worksheets.Count = 1 Then Exit Function
    
    If pIndex > 1 Then
        If pIndex <= pWB.Worksheets.Count Then
            ws.Move After:=pWB.Worksheets(pIndex)
        Else
            ws.Move After:=pWB.Worksheets(pWB.Worksheets.Count)
        End If
    Else
        ws.Move Before:=pWB.Worksheets(1)
    End If
End Function

Public Sub DeleteSheet(ByVal pWB As Workbook, ByVal pWS As Worksheet)
    Dim oldDisplayAlerts As Boolean
    
    If pWB Is Nothing Then Exit Sub
    If pWS Is Nothing Then Exit Sub
    If pWS.Parent Is Nothing Then Exit Sub
    If Not pWS.Parent Is pWB Then Exit Sub
    If pWB.Worksheets.Count <= 1 Then Exit Sub
    
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    On Error GoTo CleanUp
    pWS.Delete
    
CleanUp:
    Application.DisplayAlerts = oldDisplayAlerts
End Sub

Public Sub SortSheetsAlphabetically(ByVal pWB As Workbook)
    Dim i As Long
    Dim bSorted As Boolean
    
    If pWB Is Nothing Then Exit Sub
    If pWB.Sheets.Count < 2 Then Exit Sub
    
    ' Bubble sort
    Do
        bSorted = True
        
        For i = pWB.Sheets.Count - 1 To 1 Step -1
            If StrComp(pWB.Sheets(i).Name, pWB.Sheets(i + 1).Name, vbTextCompare) > 0 Then
                pWB.Sheets(i).Move After:=pWB.Sheets(i + 1)
                bSorted = False
            End If
        Next i
    Loop Until bSorted
End Sub

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

Public Function FindSheet(ByVal pWB As Workbook, ByVal pName As String) As Worksheet
    On Error Resume Next
    Set FindSheet = pWB.Worksheets(pName)
    On Error GoTo 0
End Function

