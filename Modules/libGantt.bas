Attribute VB_Name = "libGantt"
Option Explicit

Public Sub SetGanttColor()
    Dim iRow As Long
    Dim iStartCol As Long
    Dim iDurationCol As Long
    Dim iColorCol As Long
    Dim iLabelCol As Long
    
    ForceFindExtents
    
    iStartCol = Find_Column("Start")
    iDurationCol = Find_Column("Duration")
    iColorCol = Find_Column("Colour")
    iLabelCol = Find_Column("Label")
    
    If (iStartCol = -1) Or (iDurationCol = -1) Or (iColorCol = -1) Or (iLabelCol = -1) Then
        MsgBox ("One or more missing Columns.  Need: START, DURATION, COLOUR & LABEL")
        
        Exit Sub
    End If
    
    ActiveSheet.ChartObjects("Gantt").Activate
    ActiveChart.FullSeriesCollection(2).Select
    
    For iRow = 2 To FLastRow
        ActiveChart.FullSeriesCollection(1).Points(iRow - 1).Select
        With Selection.Format.Fill
            .Visible = msoFalse
        End With
        
        ActiveChart.FullSeriesCollection(2).Points(iRow - 1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = Cells(iRow, iColorCol).Interior.Color
            .Transparency = 0
            .Solid
        End With
    Next iRow
    
    Cells(1, 1).Select
End Sub
