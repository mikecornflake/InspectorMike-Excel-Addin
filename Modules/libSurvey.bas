Attribute VB_Name = "libSurvey"


' Uses LibraryInterpolation.InterpolateByDate
Sub Interpolate_Nav_To_3_Sec()
    Dim iDateCol As Long, iTimeCol As Long, iDateTimeCol As Long
    
    Dim dtStart As Date, dtEnd As Date
    Dim dtOneSec As Date
    Dim dtCurr As Date
    
    Dim iRow As Long, iCol As Long
    Dim iStartRow As Long, iEndRow As Long
    Dim iNewRow As Long, iNewRows As Long
    Dim iIntervalAsSec As Long
    Dim sTemp As String
    Dim iRowsAdded As Long
    
    Dim dStart As Double, dEnd As Double, dCurr As Double
    
    ForceFindExtents
    
    'Find the Date/Time columns
    iDateCol = Find_Column("Date")
    iTimeCol = Find_Column("Time")
    iDateTimeCol = FindFirstColumn(Array("Date Time", "Survey Data.Clock", "DateTime"))
    
    If (iDateTimeCol = -1) Then
        If (iDateCol = -1) Or (iTimeCol = -1) Then
            MsgBox "Missing Date Time columns"
            Exit Sub
        End If
    End If
    
    ' Constants
    dtOneSec = #12:00:01 AM#
        
    ' Go to the end of the sheet, then work back towards the start
    iRow = FLastRow - 1
    iRowsAdded = 0
    
    If MsgBox("About to commence interpolation of survey records" & vbCrLf & _
              "to maximum interval of 3 seconds." & vbCrLf & vbCrLf & _
              "This may take a while." & vbCrLf & vbCrLf & _
              "Please do not use Excel while Macro is running.", vbOKCancel) = vbOK Then
        Application.ScreenUpdating = False
        
        While (iRow >= 2)
            Cells(iRow, 1).Select
            
            Application.StatusBar = "Remaining: " & iRow
            
            If (iDateTimeCol <> -1) Then
                dtStart = Cells(iRow, iDateTimeCol).Value
                dtEnd = Cells(iRow + 1, iDateTimeCol).Value
            Else
                dtStart = Cells(iRow, iDateCol).Value + Cells(iRow, iTimeCol).Value
                dtEnd = Cells(iRow + 1, iDateCol).Value + Cells(iRow + 1, iTimeCol).Value
            End If
            
            iIntervalAsSec = DateDiff("s", dtStart, dtEnd)
            
            ' Only do the processing if there's not a LARGE GAP (possible ROV pause)
            ' Or if the gap is > than required
            If (Abs(iIntervalAsSec) < 60) And (Abs(iIntervalAsSec) > 3) Then
                iNewRows = Int(((Abs(iIntervalAsSec) - 1) / 3))
                
                For iNewRow = 1 To iNewRows
                    Rows(iRow + 1).Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    iRowsAdded = iRowsAdded + 1
                Next iNewRow
                
                iStartRow = iRow
                iEndRow = iRow + iNewRows + 1
                
                For iNewRow = 1 To iNewRows
                    Cells(iRow + iNewRow, 1).Select
                    
                    If dtStart < dtEnd Then
                        dtCurr = dtStart + iNewRow * 3 * dtOneSec
                    Else
                        dtCurr = dtStart - iNewRow * 3 * dtOneSec
                    End If
                    
                    If (iDateTimeCol <> -1) Then
                        Cells(iRow + iNewRow, iDateTimeCol).Value = dtCurr
                    Else
                        Cells(iRow + iNewRow, iDateCol).Value = Int(dtCurr)
                        Cells(iRow + iNewRow, iTimeCol).Value = Format(dtCurr, "HH:MM:SS")
                    End If
                    
                    For iCol = 1 To FLastColumn
                        If (iCol <> iDateTimeCol) And (iCol <> iDateCol) And (iCol <> iTimeCol) Then
                            sTemp = Cells(iStartRow, iCol).Value
                            If Text_IsNumber(sTemp) Then
                                dStart = Cells(iStartRow, iCol).Value
                                dEnd = Cells(iEndRow, iCol).Value
                                
                                Cells(iRow + iNewRow, iCol).Value = InterpolateByDate(dtStart, dtEnd, dtCurr, dStart, dEnd)
                            Else
                                Cells(iRow + iNewRow, iCol).Value = sTemp
                            End If
                        End If
                    Next iCol
                Next iNewRow
            End If
            
            iRow = iRow - 1
        Wend
        
        MsgBox "Completed Interpolation!  " & iRowsAdded & " rows added."
        
        Application.ScreenUpdating = True
    End If
    
    Cells(2, 1).Select
    Application.StatusBar = ""
End Sub

Public Sub BasicTidyAndFormatColumns()
    Dim iCol As Long
    
    Call BasicTidy(ActiveSheet, False)
    
    Call FormatColumnByNames(Array("Date", "#Date", "Start Date", "End Date"), "dd/mm/yyyy")
    Call FormatColumnByNames(Array("Time", "#Time", "Start Time", "Start Time (Local)", "End Time", "End Time (Local)"), "HH:mm:ss")
    Call FormatColumnByNames(Array("DateTime", "Date Time", "#Date Time", "#DateTime", "Survey Data.Clock", "Event.Start Clock", "Clock"), "dd/mm/yyyy HH:mm:ss")
    Call FormatColumnByNames(Array("Event.End Clock", "Start DateTime", "Start DateTime (Local)", "End DateTime", "End DateTime (Local)"), "dd/mm/yyyy HH:mm:ss")
    
    Call FormatColumnByNames(Array("KP", "Survey - Pipeline.KP"), "0.0000")
    Call FormatColumnByNames(Array("Easting", "Eastings", "Survey - Standard.Easting"), "0.00")
    Call FormatColumnByNames(Array("Northing", "Northings", "Survey - Standard.Northing"), "0.00")
    Call FormatColumnByNames(Array("Depth", "Depth (m)", "Depth(m)", "Survey - Standard.Depth"), "0.0")
    Call FormatColumnByNames(Array("Elevation", "Elevation (m)", "Elevation(m)", "Survey - Standard.Elevation"), "0.00")
    
    Call FormatColumnByNames(Array("Heading", "Other Fields.Heading"), "0.0")
    
    Call FormatColumnByName("Pitch", "0.0")
    Call FormatColumnByName("Roll", "0.0")
    Call FormatColumnByName("LSH", "0.00")
    Call FormatColumnByName("RSH", "0.00")
    Call FormatColumnByNames(Array("LSB", "Survey - Pipeline.Left", "PL - Profile.Left Seabed"), "0.0")
    Call FormatColumnByNames(Array("RSB", "Survey - Pipeline.Right", "PL - Profile.Right Seabed"), "0.0")
    Call FormatColumnByNames(Array("TOP", "Survey - Pipeline.ToP", "PL - Profile.Top of Pipe"), "0.0")
    Call FormatColumnByNames(Array("BOP", "Survey - Pipeline.BoP", "PL - Profile.Bottom of Pipe"), "0.0")
    Call FormatColumnByName("Salinity", "0.0")
    Call FormatColumnByName("Velocity", "0.000")
    
    Call FormatColumnByNames(Array("CP", "CP reading", "CP Readings"), "0.000")
    Call FormatColumnByNames(Array("Temperature", "Temp"), "0.0")
    Call FormatColumnByNames(Array("DVLDist", "Distance", "Survey - Pipeline.Distance"), "0.000")
    Call FormatColumnByNames(Array("DCC", "DOL", "Offset", "Survey - Pipeline.Offset"), "0.00")
    
    Cells(2, 1).Select
End Sub
