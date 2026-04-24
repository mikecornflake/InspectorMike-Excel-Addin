Attribute VB_Name = "libDate"
Option Explicit


' Converts a full timestamp string (YYYYMMDDHHMMSS) to a Date
Public Function DateTime_FromYYYYMMDDHHMMSS(ByVal pText As String) As Date
    Dim sDate As String, sTime As String
    
    sDate = Mid(pText, 1, 8)
    sTime = Mid(pText, 9, 6)
    
    DateTime_FromYYYYMMDDHHMMSS = Date_FromYYYYMMDD(sDate) + Time_FromHHMMSS(sTime)
End Function

' Converts a date string (YYYYMMDD) to a Date
Public Function Date_FromYYYYMMDD(ByVal pText As String) As Date
    Dim sY As String, sM As String, sd As String
    sY = Mid(pText, 1, 4)
    sM = Mid(pText, 5, 2)
    sd = Mid(pText, 7, 2)
    
    Date_FromYYYYMMDD = DateSerial(sY, sM, sd)
End Function

' Converts a time string (HHMMSS) to a Date (time portion only)
Public Function Time_FromHHMMSS(ByVal pText As String) As Date
    Dim sH As String, sM As String, sS As String
    
    sH = Mid(pText, 1, 2)
    sM = Mid(pText, 3, 2)
    sS = Mid(pText, 5, 2)
    
    Time_FromHHMMSS = TimeSerial(sH, sM, sS)
End Function

' Formats a Date into DD/MM/YYYY string
Public Function Date_ToDMY(ByVal pDate As Date) As String
    Date_ToDMY = Format(pDate, "DD/MM/YYYY")
End Function

' Formats a Date into HH:MM:SS string
Public Function Time_ToHMS(ByVal pTime As Date) As String
    Time_ToHMS = Format(pTime, "HH:MM:SS")
End Function

' Converts a UNIX timestamp string to a Date
Public Function DateTime_FromUnixTime(ByVal pUnixTime As String) As Date
    Dim dUnixTime As Double
    Dim iLen As Long

    On Error GoTo ErrHandler

    dUnixTime = CDbl(pUnixTime)
    iLen = Len(pUnixTime)

    Select Case iLen
        Case 10 ' Seconds since epoch
            DateTime_FromUnixTime = DateAdd("s", dUnixTime, DateSerial(1970, 1, 1))
        Case 13 ' Milliseconds since epoch
            DateTime_FromUnixTime = dUnixTime / 86400000 + DateSerial(1970, 1, 1)
        Case Else
            If dUnixTime = 0 Then
                DateTime_FromUnixTime = DateSerial(1970, 1, 1)
            Else
                DateTime_FromUnixTime = 0
            End If
    End Select
    
    Exit Function

ErrHandler:
    DateTime_FromUnixTime = 0
End Function

' Converts a Date to a time string, optionally including milliseconds
Public Function Time_ToText(ByVal pTime As Date, Optional ByVal pIncludeMilliseconds As Boolean = False) As String
    Dim sTemp As String
    Dim ms As Long
    Const imSecPerDay As Long = 86400000

    sTemp = Format(pTime, "HH:mm:ss")

    If pIncludeMilliseconds Then
        ms = Round((pTime - Fix(pTime)) * imSecPerDay)
        If ms > 0 Then sTemp = sTemp & "." & ms
    End If

    Time_ToText = sTemp
End Function

' Converts a time string (with optional milliseconds) to a Date
Public Function Time_FromText(ByVal pText As String) As Date
    Dim sTime As String, sMillisec As String
    Dim dMillisec As Double
    Dim imSecPerDay As Long
    
    imSecPerDay = CLng(24) * 60 * 60 * 1000
    dMillisec = 0
    sTime = pText

    ' Extract milliseconds if present
    If InStr(pText, ".") > 0 Then
        sTime = Text_Between(pText, "", ".")
        sMillisec = Text_Between(pText, ".", "")
        dMillisec = CDbl(sMillisec) / imSecPerDay
    End If
    
    Time_FromText = TimeValue(sTime) + dMillisec
End Function

' Converts date and time strings to a full Date value
Public Function DateTime_FromText(ByVal pDateText As String, ByVal pTimeText As String) As Date
    Dim iT As Long
    Dim dtDate As Date, dtTime As Date
    Dim sTemp As String

    ' Handle ISO format with "T" separator
    iT = InStr(pDateText, "T")
    If iT <> 0 Then
        sTemp = Text_Replace(pDateText, "T", " ")
    Else
        sTemp = pDateText
    End If

    dtDate = DateValue(sTemp)

    If pTimeText <> "" Then
        dtTime = Time_FromText(pTimeText)
    Else
        dtTime = TimeValue(sTemp)
    End If

    DateTime_FromText = dtDate + dtTime
End Function
