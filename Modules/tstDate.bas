Attribute VB_Name = "tstDate"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass
'            - Tests by copilot/ChatGPT-5

Option Explicit
Option Private Module

' Runs a suite of assertions to validate date/time conversion functions
Public Sub Test_LibraryDate()
    ActiveTestModule = "libDate"
    
    ' Expected date/time value for comparison
    Dim dtExpected As Date
    dtExpected = DateSerial(2025, 8, 16) + TimeSerial(1, 7, 0)
    
    ' Test full timestamp conversion from string
    Call AssertEqual("DateTime_FromYYYYMMDDHHMMSS - full timestamp", dtExpected, DateTime_FromYYYYMMDDHHMMSS("20250816010700"))
    
    ' Test date-only conversion
    Call AssertEqual("Date_FromYYYYMMDD - basic date", DateSerial(2025, 8, 16), Date_FromYYYYMMDD("20250816"))
    
    ' Test time-only conversion
    Call AssertEqual("Time_FromHHMMSS - basic time", TimeSerial(1, 7, 0), Time_FromHHMMSS("010700"))
    
    ' Test formatting date to DD/MM/YYYY
    Call AssertEqual("Date_ToDMY - format check", "16/08/2025", Date_ToDMY(DateSerial(2025, 8, 16)))
    
    ' Test formatting time to HH:MM:SS
    Call AssertEqual("Time_ToHMS - format check", "01:07:00", Time_ToHMS(TimeSerial(1, 7, 0)))
    
    ' Test UNIX timestamp conversion (10-digit and 13-digit)
    Call AssertEqual("DateTime_FromUnixTime - 10-digit", DateSerial(1970, 1, 1), DateTime_FromUnixTime("0"))
    Call AssertEqual("DateTime_FromUnixTime - 13-digit", DateSerial(1970, 1, 1), DateTime_FromUnixTime("0"))

    ' Create a time with fractional seconds
    Dim dt As Date
    dt = TimeSerial(1, 7, 0) + 0.00001
    
    ' Test time-to-string conversion with and without milliseconds
    Call AssertEqual("Time_ToText - no milliseconds 1", "01:07:00", Time_ToText(TimeSerial(1, 7, 0), False))
    Call AssertEqual("Time_ToText - no milliseconds 2", "01:07:01", Time_ToText(dt, False))
    Call AssertTrue("Time_ToText - with milliseconds", InStr(Time_ToText(dt, True), ".") > 0)

    ' Test string-to-time conversion
    Call AssertEqual("Time_FromText - basic", TimeSerial(1, 7, 0), Time_FromText("01:07:00"))
    Call AssertTrue("Time_FromText - with milliseconds", Time_FromText("01:07:00.500") > TimeSerial(1, 7, 0))
    
    ' Test string-to-date conversion with separate and combined inputs
    dtExpected = DateSerial(2025, 8, 16) + TimeSerial(1, 7, 0)
    Call AssertEqual("DateTime_FromText - separate date/time", dtExpected, DateTime_FromText("2025-08-16", "01:07:00"))
    Call AssertEqual("DateTime_FromText - combined ISO", dtExpected, DateTime_FromText("2025-08-16T01:07:00", ""))
    Call AssertEqual("DateTime_FromText - combined with space", dtExpected, DateTime_FromText("2025-08-16 01:07:00", ""))
End Sub

