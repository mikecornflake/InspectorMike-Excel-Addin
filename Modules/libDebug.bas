Attribute VB_Name = "libDebug"
Option Explicit
Option Private Module

'===========================================================
' libDebgug
'
' 7 May 2026
'  Create a temp sheet xe.debug if needed
'  Export log entries via debug_log
'  Optional cache writes with debug_BeginUpdate and debug_EndUpdate
'  Optional Indent via debug_IncIndent / debug_DecIndent
'  Reset log entried with debug_Clear
'===========================================================


Private mIndent As Long   ' Tracks current indentation level (in spaces)
Private mNextRow As Long  ' Cache the next row

' Update cache
Private mUpdateDepth As Long
Private mRows As Collection

'===========================================================
' Main Debug Logger
'===========================================================
Public Sub debug_Log(wsName As String, msg As String)
    Dim indentText As String
    indentText = String(mIndent, " ")

    If mUpdateDepth > 0 Then
        Dim row(1 To 3) As Variant
        row(1) = Now
        row(2) = wsName
        row(3) = indentText & msg
        mRows.Add row
        Exit Sub
    End If

    ' normal immediate write...
    Dim ws As Worksheet
    Set ws = GetDebugSheet()

    If mNextRow = 0 Then
        mNextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    End If

    With ws
        .Cells(mNextRow, 1).Value = Now
        .Cells(mNextRow, 2).Value = wsName
        .Cells(mNextRow, 3).Value = indentText & msg
    End With

    mNextRow = mNextRow + 1
End Sub

'===========================================================
' Increase indentation by 2 spaces
'===========================================================
Public Sub debug_IncIndent()
    mIndent = mIndent + 2
End Sub

'===========================================================
' Decrease indentation by 2 spaces (not below zero)
'===========================================================
Public Sub debug_DecIndent()
    mIndent = mIndent - 2
    If mIndent < 0 Then mIndent = 0
End Sub

'===========================================================
' Clear debug log (keep header), ensure sheet exists/visible,
' reset indentation
'===========================================================
Public Sub debug_Clear()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("xe.debug")
    On Error GoTo 0
    
    mNextRow = 0
    mIndent = 0    ' reset the indent
    
    If Not ws Is Nothing Then
        ws.Rows("2:" & ws.Rows.Count).ClearContents
        mNextRow = 2   ' next write starts at row 2
    End If

End Sub

'===========================================================
' Start caching the debug_logs instead of instant write
'===========================================================
Public Sub debug_BeginUpdate()
    mUpdateDepth = mUpdateDepth + 1
    If mUpdateDepth = 1 Then
        Set mRows = New Collection
    End If
End Sub

'===========================================================
' End caching the debug_logs, output cache in a single hit
'===========================================================
Public Sub debug_EndUpdate()
    If mUpdateDepth = 0 Then Exit Sub

    mUpdateDepth = mUpdateDepth - 1

    If mUpdateDepth = 0 Then
        FlushCollectionBuffer
    End If
End Sub

'===========================================================
' Helper: Output Collection of Rows to WorkSheet
'===========================================================
Private Sub FlushCollectionBuffer()
    If mRows Is Nothing Or mRows.Count = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = GetDebugSheet()

    If mNextRow = 0 Then
        mNextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    End If

    Dim arr() As Variant
    ReDim arr(1 To mRows.Count, 1 To 3)

    Dim i As Long
    For i = 1 To mRows.Count
        arr(i, 1) = mRows(i)(1)
        arr(i, 2) = mRows(i)(2)
        arr(i, 3) = mRows(i)(3)
    Next i

    ws.Cells(mNextRow, 1).Resize(mRows.Count, 3).Value = arr
    mNextRow = mNextRow + mRows.Count

    Set mRows = Nothing
End Sub

'===========================================================
' Helper: Pretty format the debug log
'===========================================================
Private Sub debug_Format()
    Dim ws As Worksheet
    
    Dim bUpdating As Boolean, bEventsEnabled As Boolean
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("xe.debug")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        bUpdating = Application.ScreenUpdating
        bEventsEnabled = Application.EnableEvents
        
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
        ' Apply autofilter
        ws.Range("A1:C1").AutoFilter
    
        ' Freeze header row
        With ws
            .Activate
            .Range("A2").Select
            ActiveWindow.FreezePanes = True
        End With
    
        ' Header colour
        With ws.Range("A1:C1").Interior
            .Pattern = xlSolid
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.8
        End With
        
        ' Set the format for the date row
        ws.Columns(1).NumberFormat = "dd/mm/yyyy HH:mm:ss"
    
        ' Resize columns
        Columns("A:A").ColumnWidth = 18.5
        Columns("B:B").ColumnWidth = 18.5
        Columns("C:C").ColumnWidth = 75
    
        Application.ScreenUpdating = bUpdating
        Application.EnableEvents = bEventsEnabled
    End If
End Sub

'===========================================================
' Helper: Get or create xe.debug sheet
'===========================================================
Private Function GetDebugSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("xe.debug")
    On Error GoTo 0
    
    ' Create if missing
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "xe.debug"
        
        ' Add headers
        ws.Range("A1").Value = "DateTime"
        ws.Range("B1").Value = "Worksheet"
        ws.Range("C1").Value = "Description"
        
        ws.Rows(1).Font.Bold = True
        
        mNextRow = 2   ' next write starts at row 2
        
        ' Format the sheet
        debug_Format
    End If
    
    ' Unhide if hidden
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If
    
    ' Colour the Tab yellow
    With ws.Tab
        .Color = 65535
        .TintAndShade = 0
    End With
    
    Set GetDebugSheet = ws
End Function
