Attribute VB_Name = "tstInterpolation"
Option Explicit
Option Private Module

Public Sub Test_libInterpolation()
    ActiveTestModule = "libInterpolation"

    Call Test_InterpolateByDouble
    Call Test_InterpolateByDouble_Errors
    
    Call Test_InterpolateByDate
    Call Test_InterpolateByDate_Errors
    
    Call Test_InterpolateByDoubleClamped
    Call Test_InterpolateByDoubleClamped_Errors

    Call Test_InterpolateByDateClamped
    Call Test_InterpolateByDateClamped_Errors

End Sub

Private Sub Test_InterpolateByDouble()
    ' Basic midpoint
    Call AssertEqual("InterpolateByDouble - midpoint", 50#, InterpolateByDouble(0#, 100#, 50#, 0#, 100#))

    ' Quarter / three-quarter points
    Call AssertEqual("InterpolateByDouble - quarter", 25#, InterpolateByDouble(0#, 100#, 25#, 0#, 100#))
    Call AssertEqual("InterpolateByDouble - three quarter", 75#, InterpolateByDouble(0#, 100#, 75#, 0#, 100#))

    ' Non-zero start/end range
    Call AssertEqual("InterpolateByDouble - non-zero x range", 15#, InterpolateByDouble(10#, 20#, 15#, 10#, 20#))

    ' Output range not matching input range
    Call AssertEqual("InterpolateByDouble - scaled output", 150#, InterpolateByDouble(0#, 10#, 5#, 100#, 200#))

    ' Descending Y values
    Call AssertEqual("InterpolateByDouble - descending y", 50#, InterpolateByDouble(0#, 100#, 50#, 100#, 0#))

    ' Extrapolation before start
    Call AssertEqual("InterpolateByDouble - extrapolate before", -50#, InterpolateByDouble(0#, 100#, -50#, 0#, 100#))

    ' Extrapolation after end
    Call AssertEqual("InterpolateByDouble - extrapolate after", 150#, InterpolateByDouble(0#, 100#, 150#, 0#, 100#))

    ' Decimal values
    Call AssertEqual("InterpolateByDouble - decimals", 1.5, InterpolateByDouble(0#, 2#, 1#, 1#, 2#))

    ' Negative x values
    Call AssertEqual("InterpolateByDouble - negative x range", 50#, InterpolateByDouble(-100#, 100#, 0#, 0#, 100#))

    ' Negative y values
    Call AssertEqual("InterpolateByDouble - negative y range", 0#, InterpolateByDouble(0#, 100#, 50#, -100#, 100#))

    ' Reversed x range currently works mathematically
    Call AssertEqual("InterpolateByDouble - reversed x range", 50#, InterpolateByDouble(100#, 0#, 50#, 0#, 100#))
End Sub

Private Sub Test_InterpolateByDate()
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim dtCurr As Date

    dtStart = #1/1/2025#
    dtEnd = #1/11/2025#

    ' Start point
    Call AssertEqual("InterpolateByDate - start", 0#, InterpolateByDate(dtStart, dtEnd, dtStart, 0#, 100#))

    ' End point
    Call AssertEqual("InterpolateByDate - end", 100#, InterpolateByDate(dtStart, dtEnd, dtEnd, 0#, 100#))

    ' Midpoint
    dtCurr = #1/6/2025#
    Call AssertEqual("InterpolateByDate - midpoint", 50#, InterpolateByDate(dtStart, dtEnd, dtCurr, 0#, 100#))

    ' Quarter-ish point: 2.5 days via time fraction
    dtCurr = dtStart + 2.5
    Call AssertEqual("InterpolateByDate - quarter with time", 25#, InterpolateByDate(dtStart, dtEnd, dtCurr, 0#, 100#))

    ' Scaled output
    dtCurr = #1/6/2025#
    Call AssertEqual("InterpolateByDate - scaled output", 150#, InterpolateByDate(dtStart, dtEnd, dtCurr, 100#, 200#))

    ' Descending Y values
    Call AssertEqual("InterpolateByDate - descending y", 50#, InterpolateByDate(dtStart, dtEnd, dtCurr, 100#, 0#))

    ' Extrapolation before start
    dtCurr = #12/27/2024#
    Call AssertEqual("InterpolateByDate - extrapolate before", -50#, InterpolateByDate(dtStart, dtEnd, dtCurr, 0#, 100#))

    ' Extrapolation after end
    dtCurr = #1/16/2025#
    Call AssertEqual("InterpolateByDate - extrapolate after", 150#, InterpolateByDate(dtStart, dtEnd, dtCurr, 0#, 100#))

    ' Reversed date range currently works mathematically
    Call AssertEqual("InterpolateByDate - reversed date range", 50#, InterpolateByDate(dtEnd, dtStart, #1/6/2025#, 0#, 100#))
End Sub

Private Sub Test_InterpolateByDouble_Errors()
    Dim lErr As Long

    On Error Resume Next
    Call InterpolateByDouble(10#, 10#, 10#, 0#, 100#)
    lErr = Err.Number
    Err.Clear
    On Error GoTo 0

    Call AssertRaises( _
        "InterpolateByDouble - zero range raises error", _
        ERR_INTERPOLATION_ZERO_RANGE, _
        lErr)
End Sub

Private Sub Test_InterpolateByDate_Errors()
    Dim lErr As Long
    Dim dtStart As Date

    dtStart = #1/1/2025#

    On Error Resume Next
    Call InterpolateByDate(dtStart, dtStart, dtStart, 0#, 100#)
    lErr = Err.Number
    Err.Clear
    On Error GoTo 0

    Call AssertRaises( _
        "InterpolateByDate - zero range raises error", _
        ERR_INTERPOLATION_ZERO_RANGE, _
        lErr)
End Sub

Private Sub Test_InterpolateByDoubleClamped()
    Call AssertEqual("InterpolateByDoubleClamped - midpoint", 50#, InterpolateByDoubleClamped(0#, 100#, 50#, 0#, 100#))

    Call AssertEqual("InterpolateByDoubleClamped - before start clamps", 0#, InterpolateByDoubleClamped(0#, 100#, -50#, 0#, 100#))

    Call AssertEqual("InterpolateByDoubleClamped - after end clamps", 100#, InterpolateByDoubleClamped(0#, 100#, 150#, 0#, 100#))

    Call AssertEqual("InterpolateByDoubleClamped - descending y before start clamps", 100#, InterpolateByDoubleClamped(0#, 100#, -50#, 100#, 0#))

    Call AssertEqual("InterpolateByDoubleClamped - descending y after end clamps", 0#, InterpolateByDoubleClamped(0#, 100#, 150#, 100#, 0#))

    Call AssertEqual("InterpolateByDoubleClamped - reversed x midpoint", 50#, InterpolateByDoubleClamped(100#, 0#, 50#, 0#, 100#))

    Call AssertEqual("InterpolateByDoubleClamped - reversed x before start clamps", 0#, InterpolateByDoubleClamped(100#, 0#, 150#, 0#, 100#))

    Call AssertEqual("InterpolateByDoubleClamped - reversed x after end clamps", 100#, InterpolateByDoubleClamped(100#, 0#, -50#, 0#, 100#))
End Sub

Private Sub Test_InterpolateByDoubleClamped_Errors()
    Dim lErr As Long

    On Error Resume Next
    Call InterpolateByDoubleClamped(10#, 10#, 10#, 0#, 100#)
    lErr = Err.Number
    Err.Clear
    On Error GoTo 0

    Call AssertRaises( _
        "InterpolateByDoubleClamped - zero range raises error", _
        ERR_INTERPOLATION_ZERO_RANGE, _
        lErr)
End Sub

Private Sub Test_InterpolateByDateClamped()
    Dim dtStart As Date
    Dim dtEnd As Date

    dtStart = #1/1/2025#
    dtEnd = #1/11/2025#

    Call AssertEqual("InterpolateByDateClamped - midpoint", 50#, InterpolateByDateClamped(dtStart, dtEnd, #1/6/2025#, 0#, 100#))

    Call AssertEqual("InterpolateByDateClamped - before start clamps", 0#, InterpolateByDateClamped(dtStart, dtEnd, #12/27/2024#, 0#, 100#))

    Call AssertEqual("InterpolateByDateClamped - after end clamps", 100#, InterpolateByDateClamped(dtStart, dtEnd, #1/16/2025#, 0#, 100#))

    Call AssertEqual("InterpolateByDateClamped - descending y before start clamps", 100#, InterpolateByDateClamped(dtStart, dtEnd, #12/27/2024#, 100#, 0#))

    Call AssertEqual("InterpolateByDateClamped - descending y after end clamps", 0#, InterpolateByDateClamped(dtStart, dtEnd, #1/16/2025#, 100#, 0#))

    Call AssertEqual("InterpolateByDateClamped - reversed dates midpoint", 50#, InterpolateByDateClamped(dtEnd, dtStart, #1/6/2025#, 0#, 100#))

    Call AssertEqual("InterpolateByDateClamped - reversed dates before start clamps", 0#, InterpolateByDateClamped(dtEnd, dtStart, #1/16/2025#, 0#, 100#))

    Call AssertEqual("InterpolateByDateClamped - reversed dates after end clamps", 100#, InterpolateByDateClamped(dtEnd, dtStart, #12/27/2024#, 0#, 100#))
End Sub

Private Sub Test_InterpolateByDateClamped_Errors()
    Dim lErr As Long
    Dim dtStart As Date

    dtStart = #1/1/2025#

    On Error Resume Next
    Call InterpolateByDateClamped(dtStart, dtStart, dtStart, 0#, 100#)
    lErr = Err.Number
    Err.Clear
    On Error GoTo 0

    Call AssertRaises( _
        "InterpolateByDateClamped - zero range raises error", _
        ERR_INTERPOLATION_ZERO_RANGE, _
        lErr)
End Sub
