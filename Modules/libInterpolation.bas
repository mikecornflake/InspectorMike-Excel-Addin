Attribute VB_Name = "libInterpolation"
Option Explicit
Option Private Module

Private Sub Interpolation_RaiseZeroRangeError(ByVal pRoutineName As String)
    Err.Raise _
        Number:=ERR_INTERPOLATION_ZERO_RANGE, _
        source:="libInterpolation." & pRoutineName, _
        Description:="Start and end values cannot be equal."
End Sub

' Routines named InterpolateXXX actually allow for extrapolation
' Up to caller to determine allowable extrapolation range

Private Function InterpolateCore(ByVal dStart As Double, ByVal dEnd As Double, _
                                 ByVal dCurr As Double, _
                                 ByVal dY1 As Double, ByVal dY2 As Double, _
                                 ByVal pRoutineName As String) As Double
                                 
    If dEnd = dStart Then
        Call Interpolation_RaiseZeroRangeError(pRoutineName)
    End If
    
    Dim dPercent As Double

    dPercent = (dCurr - dStart) / (dEnd - dStart)
    InterpolateCore = dY1 + (dPercent * (dY2 - dY1))
End Function

Public Function InterpolateByDate(ByVal dtStart As Date, ByVal dtEnd As Date, _
                                  ByVal dtCurr As Date, _
                                  ByVal dY1 As Double, ByVal dY2 As Double) As Double
    
    InterpolateByDate = InterpolateCore(CDbl(dtStart), CDbl(dtEnd), CDbl(dtCurr), dY1, dY2, "InterpolateByDate")
End Function

Public Function InterpolateByDouble(ByVal dStart As Double, ByVal dEnd As Double, _
                                    ByVal dCurr As Double, _
                                    ByVal dY1 As Double, ByVal dY2 As Double) As Double
    
    InterpolateByDouble = InterpolateCore(dStart, dEnd, dCurr, dY1, dY2, "InterpolateByDouble")
End Function

' The XXXClamped functions do not allow extrapolation
' Instead the move the extrapolation point to either
' the Start or the End value, depending on which
' side it's outside of

Public Function InterpolateByDoubleClamped( _
                    ByVal dStart As Double, ByVal dEnd As Double, _
                    ByVal dCurr As Double, _
                    ByVal dY1 As Double, ByVal dY2 As Double) As Double


    If dEnd = dStart Then
        Call Interpolation_RaiseZeroRangeError("InterpolateByDoubleClamped")
    End If

    If dStart < dEnd Then
        If dCurr < dStart Then dCurr = dStart
        If dCurr > dEnd Then dCurr = dEnd
    Else
        If dCurr > dStart Then dCurr = dStart
        If dCurr < dEnd Then dCurr = dEnd
    End If

    InterpolateByDoubleClamped = InterpolateCore( _
        dStart, dEnd, dCurr, dY1, dY2, "InterpolateByDoubleClamped")
End Function

Public Function InterpolateByDateClamped( _
                    ByVal dtStart As Date, ByVal dtEnd As Date, _
                    ByVal dtCurr As Date, _
                    ByVal dY1 As Double, ByVal dY2 As Double) As Double

    If dtEnd = dtStart Then
        Call Interpolation_RaiseZeroRangeError("InterpolateByDateClamped")
    End If

    If dtStart < dtEnd Then
        If dtCurr < dtStart Then dtCurr = dtStart
        If dtCurr > dtEnd Then dtCurr = dtEnd
    Else
        If dtCurr > dtStart Then dtCurr = dtStart
        If dtCurr < dtEnd Then dtCurr = dtEnd
    End If

    InterpolateByDateClamped = InterpolateCore( _
        CDbl(dtStart), CDbl(dtEnd), CDbl(dtCurr), dY1, dY2, "InterpolateByDateClamped")
End Function

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
                If oCell.row < iFirstRow Then
                    iFirstRow = oCell.row
                End If
                  
                If oCell.row > iLastRow Then
                    iLastRow = oCell.row
                End If
            End If
        Next oCell
        
        If iFirstRow <> 2147483647 Then
            vStart = Cells(iFirstRow, iCol)
            vEnd = Cells(iLastRow, iCol)
            
            
            For Each oCell In rngSelected
                If (oCell.Column = iCol) And (oCell.row <> iFirstRow) And (oCell.row <> iLastRow) Then
                    iRow = oCell.row - iFirstRow
                    
                    oCell.Value = vStart + iRow * ((vEnd - vStart) / (iLastRow - iFirstRow))
                    oCell.Select
                    ColorSelected
                End If
            Next oCell
        End If
    Next iCol
End Sub

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
