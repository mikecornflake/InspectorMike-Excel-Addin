Attribute VB_Name = "libInterpolation"
Function InterpolateByDate(dtStart As Date, dtEnd As Date, dtCurr As Date, dY1 As Double, dY2 As Double) As Double
    Dim dPercent As Double
    
'    If dtEnd < dtStart Then
'        Err.Raise 555, "MathRoutine.Interpolate", "Error, times in wrong order..."
'    End If
    
    dPercent = (dtCurr - dtStart) / (dtEnd - dtStart)
    InterpolateByDate = dY1 + (dPercent * (dY2 - dY1))
End Function

Function InterpolateByDouble(dStart As Double, dEnd As Double, dCurr As Double, dY1 As Double, dY2 As Double) As Double
    Dim dPercent As Double
    
    dPercent = (dCurr - dStart) / (dEnd - dStart)
    InterpolateByDouble = dY1 + (dPercent * (dY2 - dY1))
End Function

Private Sub ColorSelected()
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    With Selection.Font
        .Color = -16383844
        .TintAndShade = 0
    End With
End Sub

Private Sub InterpolateSelectedRangeByColumns()
    ' Performs a linear interpolation, filling empty cells based on values in top selected cell and bottom selected cell of each selected column
    Dim iFirstCol As Long, iLastCol As Long
    Dim iFirstRow As Long, iLastRow As Long, iRow As Long
    Dim oCell As Range, rngSelected As Range
    Dim vStart, vEnd
    Dim iCol As Long
    
    Set rngSelected = Selection
    
    iFirstCol = 2147483647
    iLastCol = -1
    
    For Each oCell In rngSelected
        If oCell.Column > iLastCol Then
            iLastCol = oCell.Column
        End If
          
        If oCell.Column < iFirstCol Then
            iFirstCol = oCell.Column
        End If
    Next oCell
    
    For iCol = iFirstCol To iLastCol
        iFirstRow = 2147483647
        iLastRow = -1
        
        For Each oCell In rngSelected
            If oCell.Column = iCol Then
                If oCell.Row < iFirstRow Then
                    iFirstRow = oCell.Row
                End If
                  
                If oCell.Row > iLastRow Then
                    iLastRow = oCell.Row
                End If
            End If
        Next oCell
        
        If iFirstRow <> 2147483647 Then
            vStart = Cells(iFirstRow, iCol)
            vEnd = Cells(iLastRow, iCol)
            
            
            For Each oCell In rngSelected
                If (oCell.Column = iCol) And (oCell.Row <> iFirstRow) And (oCell.Row <> iLastRow) Then
                    iRow = oCell.Row - iFirstRow
                    
                    oCell.Value = vStart + iRow * ((vEnd - vStart) / (iLastRow - iFirstRow))
                    oCell.Select
                    ColorSelected
                End If
            Next oCell
        End If
    Next iCol
End Sub


