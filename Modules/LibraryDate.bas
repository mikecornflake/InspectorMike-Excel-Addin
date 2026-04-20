Attribute VB_Name = "LibraryDate"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass
'            - Tests by copilot/ChatGPT-5

Option Explicit

' Runs a suite of assertions to validate date/time conversion functions
Public Sub Test_LibraryDate()
    ActiveTestModule = "LibraryDate"
    
    ' Expected date/time value for comparison
    Dim dtExpected As Date
    dtExpected = DateSerial(2025, 8, 16) + TimeSerial(1, 7, 0)
    
    ' Test full timestamp conversion from string
    Call AssertEqual("ConvertYYMMDDHHMMSSToDate - full timestamp", dtExpected, ConvertYYMMDDHHMMSSToDate("20250816010700"))
    
    ' Test date-only conversion
    Call AssertEqual("ConvertYYYYMMDDToDate - basic date", DateSerial(2025, 8, 16), ConvertYYYYMMDDToDate("20250816"))
    
    ' Test time-only conversion
    Call AssertEqual("ConvertHHMMSSToDate - basic time", TimeSerial(1, 7, 0), ConvertHHMMSSToDate("010700"))
    
    ' Test formatting date to DD/MM/YYYY
    Call AssertEqual("ConvertDateToDDMMYYYY - format check", "16/08/2025", ConvertDateToDDMMYYYY(DateSerial(2025, 8, 16)))
    
    ' Test formatting time to HH:MM:SS
    Call AssertEqual("ConvertDateToHHMMSS - format check", "01:07:00", ConvertDateToHHMMSS(TimeSerial(1, 7, 0)))
    
    ' Test UNIX timestamp conversion (10-digit and 13-digit)
    Call AssertEqual("UNIXTimeToDate - 10-digit", DateSerial(1970, 1, 1), UNIXTimeToDate("0"))
    Call AssertEqual("UNIXTimeToDate - 13-digit", DateSerial(1970, 1, 1), UNIXTimeToDate("0"))

    ' Create a time with fractional seconds
    Dim dt As Date
    dt = TimeSerial(1, 7, 0) + 0.00001
    
    ' Test time-to-string conversion with and without milliseconds
    Call AssertEqual("TimeToStr - no milliseconds 1", "01:07:00", TimeToStr(TimeSerial(1, 7, 0), False))
    Call AssertEqual("TimeToStr - no milliseconds 2", "01:07:01", TimeToStr(dt, False))
    Call AssertTrue("TimeToStr - with milliseconds", InStr(TimeToStr(dt, True), ".") > 0)

    ' Test string-to-time conversion
    Call AssertEqual("StrToTime - basic", TimeSerial(1, 7, 0), StrToTime("01:07:00"))
    Call AssertTrue("StrToTime - with milliseconds", StrToTime("01:07:00.500") > TimeSerial(1, 7, 0))
    
    ' Test string-to-date conversion with separate and combined inputs
    dtExpected = DateSerial(2025, 8, 16) + TimeSerial(1, 7, 0)
    Call AssertEqual("StrToDate - separate date/time", dtExpected, StrToDate("2025-08-16", "01:07:00"))
    Call AssertEqual("StrToDate - combined ISO", dtExpected, StrToDate("2025-08-16T01:07:00", ""))
End Sub


' Converts a full timestamp string (YYYYMMDDHHMMSS) to a Date
Public Function ConvertYYMMDDHHMMSSToDate(AInput As String) As Date
    Dim sDate As String, sTime As String
    sDate = Mid(AInput, 1, 8)
    sTime = Mid(AInput, 9, 6)
    
    ConvertYYMMDDHHMMSSToDate = ConvertYYYYMMDDToDate(sDate) + ConvertHHMMSSToDate(sTime)
End Function

' Converts a date string (YYYYMMDD) to a Date
Public Function ConvertYYYYMMDDToDate(AInput As String) As Date
    Dim sY As String, sM As String, sd As String
    sY = Mid(AInput, 1, 4)
    sM = Mid(AInput, 5, 2)
    sd = Mid(AInput, 7, 2)
    
    ConvertYYYYMMDDToDate = DateSerial(sY, sM, sd)
End Function

' Converts a time string (HHMMSS) to a Date (time portion only)
Public Function ConvertHHMMSSToDate(AInput As String) As Date
    Dim sH As String, sM As String, sS As String
    sH = Mid(AInput, 1, 2)
    sM = Mid(AInput, 3, 2)
    sS = Mid(AInput, 5, 2)
    
    ConvertHHMMSSToDate = TimeSerial(sH, sM, sS)
End Function

' Formats a Date into DD/MM/YYYY string
Public Function ConvertDateToDDMMYYYY(AInput As Date) As String
    ConvertDateToDDMMYYYY = Format(AInput, "DD/MM/YYYY")
End Function

' Formats a Date into HH:MM:SS string
Public Function ConvertDateToHHMMSS(AInput As Date) As String
    ConvertDateToHHMMSS = Format(AInput, "HH:MM:SS")
End Function

' Converts a UNIX timestamp string to a Date
Public Function UNIXTimeToDate(AUnixTime As String) As Date
    Dim dUnixTime As Double
    Dim iLen As Integer

    On Error GoTo ErrHandler

    dUnixTime = CDbl(AUnixTime)
    iLen = Len(AUnixTime)

    Select Case iLen
        Case 10 ' Seconds since epoch
            UNIXTimeToDate = DateAdd("s", dUnixTime, DateSerial(1970, 1, 1))
        Case 13 ' Milliseconds since epoch
            UNIXTimeToDate = dUnixTime / 86400000 + DateSerial(1970, 1, 1)
        Case Else
            If dUnixTime = 0 Then
                UNIXTimeToDate = DateSerial(1970, 1, 1)
            Else
                UNIXTimeToDate = 0
            End If
    End Select
    Exit Function

ErrHandler:
    UNIXTimeToDate = 0
End Function

' Converts a Date to a time string, optionally including milliseconds
Public Function TimeToStr(AInput As Date, BHasMilliSec As Boolean) As String
    Dim sTemp As String
    Dim ms As Long
    Const imSecPerDay As Long = 86400000

    sTemp = Format(AInput, "HH:mm:ss")

    If BHasMilliSec Then
        ms = Round((AInput - Fix(AInput)) * imSecPerDay)
        If ms > 0 Then sTemp = sTemp & "." & ms
    End If

    TimeToStr = sTemp
End Function

' Converts a time string (with optional milliseconds) to a Date
Public Function StrToTime(AInput As String) As Date
    Dim sTime As String, sMillisec As String
    Dim dMillisec As Double
    Dim imSecPerDay As Long
    
    imSecPerDay = CLng(24) * 60 * 60 * 1000
    dMillisec = 0
    sTime = AInput

    ' Extract milliseconds if present
    If InStr(AInput, ".") > 0 Then
        sTime = StringBetween(AInput, "", ".")
        sMillisec = StringBetween(AInput, ".", "")
        dMillisec = CDbl(sMillisec) / imSecPerDay
    End If
    
    StrToTime = TimeValue(sTime) + dMillisec
End Function

' Converts date and time strings to a full Date value
Public Function StrToDate(ADate As String, ATime As String) As Double
    Dim iT As Long
    Dim dtDate As Date, dtTime As Date
    Dim sTemp As String

    ' Handle ISO format with "T" separator
    iT = InStr(ADate, "T")
    If iT <> 0 Then
        sTemp = SwapString(ADate, "T", " ")
    Else
        sTemp = ADate
    End If

    dtDate = DateValue(sTemp)

    If ATime <> "" Then
        dtTime = StrToTime(ATime)
    Else
        dtTime = TimeValue(sTemp)
    End If

    StrToDate = dtDate + dtTime
End Function

' Reorders date components in a column to fix cross-day/month issues
Public Sub UnCrossDayMonInDateCol(ADateCol As Long)
    ForceFindExtents
    
    Dim sCol As String
    
    sCol = GetColumnLetter(ADateCol + 1)
    
    ' Insert helper column and apply corrected formula
    Columns(ADateCol + 1).Insert
    Cells(1, ADateCol + 1).Value = Cells(1, ADateCol).Value
    Cells(2, ADateCol + 1).FormulaR1C1 = "=DATE(YEAR(RC[-1]), DAY(RC[-1]), MONTH(RC[-1])) + TIME(HOUR(RC[-1]), MINUTE(RC[-1]), SECOND(RC[-1]))"
    Cells(2, ADateCol + 1).Select
    
    Selection.AutoFill Destination:=Range(sCol & "2:" & sCol & FLastRow)
    
    Columns(ADateCol + 1).Select
    Selection.Copy
    
    Cells(1, ADateCol).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Columns(ADateCol + 1).Delete
    
    Cells(1, ADateCol).Select
End Sub
