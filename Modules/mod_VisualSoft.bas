Attribute VB_Name = "mod_VisualSoft"
Option Explicit
Option Private Module

Public Type DuplicateChecks
   Code As String
   LastKP As Double
   lastRow As Long
End Type

Public Type IncidentPair
   StartCode As String
   EndCode As String
   StartKP As Double
   StartRow As Long
End Type

Public FAscendingInspection As Boolean
Public FKPThresholdForSameness As Double
Public FDuplicateEventCodeChecks() As DuplicateChecks
Public FDuplicateIncidentTypeChecks() As DuplicateChecks
Public FIncidentPairs() As IncidentPair

Private sInitialisationHack As String

Sub ShowOptions()
    InitialiseGlobalVars
    
    frmOptions.Show
End Sub

Sub InitialiseGlobalVars()
'
' This can be called multiple times, it'll only do the main code once
'

    If sInitialisationHack = vbNullString Then
        ' Place all Initialisation code here
        FAscendingInspection = True
        FKPThresholdForSameness = 0.003  ' ie Field Joints within 3m of each other will be identified as possibly the same
        
        sInitialisationHack = "Initialised"
    End If
End Sub

Function IsTidy() As Boolean
    Range("A1").Select
    IsTidy = Selection.Font.Bold
End Function

Function IsQC() As Boolean
    IsQC = Cells(1, 1).Value = "Has Issue?"
End Function

Sub TidyVWExcelExport()
    ForceFindExtents

    ' Convert all the malformed numbers back to actual numbers
    
    ' Stick a 1 on the clipboard
    Cells(FLastRow + 1, 1).Value = 1
    Cells(FLastRow + 1, 1).Select
    Selection.Copy
    
    ' Select Numbers, Paste Special, Multiple, Clear 1
    Range(Cells(2, FindColumn("Easting")), Cells(FLastRow, FindColumn("KP Length"))).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, SkipBlanks:=False, Transpose:=False
    
    Range(Cells(2, FindColumn("Contact CP")), Cells(FLastRow, FindColumn("ROV Heading"))).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, SkipBlanks:=False, Transpose:=False
    
    ' Delete the row we placed the 1 on...
    Rows(FLastRow + 1).Delete

    ' Make it all look good
    FormatActiveSheet
    
    ' Finish by selecting a sane default position
    Range("A2").Select
End Sub

Sub ProcessAnomalies()
    Dim iSheetEvent, iSheetAnomaly, i, iAnom, iAnomCol As Long
    
    ' Create the Anomaly Tab Sheet
    iSheetEvent = ActiveSheet.Index
    Sheets.Add After:=Sheets(Sheets.Count)
    iSheetAnomaly = ActiveSheet.Index
    Sheets(iSheetAnomaly).Name = "Anomalies"
    
    ' Set the Header row
    Sheets(iSheetEvent).Select
    ForceFindExtents
    iAnomCol = FindColumn("Anomaly")
    
    Rows(1).Select
    Selection.Copy
    
    Sheets(iSheetAnomaly).Select
    Rows("1:1").Select
    ActiveSheet.Paste
    
    iAnom = 2
    
    For i = 2 To FLastRow
        Sheets(iSheetEvent).Select
        
        If Cells(i, iAnomCol).Value = "Yes" Then
            ' Copy row to anomaly worksheet
            Sheets(iSheetEvent).Select
            Rows(i).Select
            Selection.Copy
            
            Sheets(iSheetAnomaly).Select
            Rows(iAnom).Select
            ActiveSheet.Paste
            iAnom = iAnom + 1
            
            ' Colour row red
            Sheets(iSheetEvent).Select
            Rows(i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 13421823
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    Sheets(iSheetAnomaly).Select
    FormatActiveSheet
    Range("A2").Select
    
    Sheets(iSheetEvent).Select
    
    ' Finish by selecting a sane default position
    Range("A2").Select
End Sub

Sub QCChecks()
    Dim i, i1, i2, i3 As Long
    Dim iDupCheck As Long
    
    ' Current Values
    Dim sIncidentType, sIncident As String
    Dim sLastSpan, sLastRockDump, sLastMattress, sLastTrench, sLastBurial, sLastInsp As String
    Dim sLocation, sComment, sAnomalyComment As String
    
    ' Columns
    Dim iIncidentType, iIncident As Long
    Dim iLastSpan, iLastRockDump, iLastMattress, iLastTrench, iLastBurial, iLastInsp As Long
    Dim iHeight, iWidth, iLength As Long
    Dim iContactCP, iCPCalibration, iAnomaly, iComment, iAnomalyComment As Long
    Dim iLocation, iClockPosition, iDistanceOff As Long
    Dim iDateCol, iTimeCol, iKPCol, iEventCodeCol As Long
    Dim iDuplicatesCol As Long
    Dim dCurrKP As Double
    Dim iDimCount As Long
    Dim iPass As Long
    
    ' Ensure the Autofilter is turned on...
    If ActiveSheet.AutoFilterMode = False Then
        Range("A1").Select
        Selection.AutoFilter
    End If
    
    ' Insert a "Pass" column, but must be sorted by time first
    ForceFindExtents
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iDateCol), Cells(FLastRow, iDateCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iTimeCol), Cells(FLastRow, iTimeCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, 1).FormulaR1C1 = "Pass"
    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    iPass = 0
    iIncident = FindColumn("Incident Code")
    For i = 2 To FLastRow
        If Cells(i, iIncident).Value = "INS" Then
            iPass = iPass + 1
        End If
        
        Cells(i, 1).Value = iPass
    Next i
    
    ' Turn Autofilter off, then back on with the new column
    Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    ' Sort by KP
    ForceFindExtents
    
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    iKPCol = FindColumn("KP")
    iEventCodeCol = FindColumn("Event Code")
    
    ' I'm no longer convinced sort order matters here
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iKPCol), Cells(FLastRow, iKPCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Insert a column to detect the duplicates that cause Coabis issues (Same Incident at Same KP)
    If FindColumn("Duplicates") = -1 Then
        ' Stop CountIf updating every time a new item is added
        Application.Calculation = xlCalculationManual
        
        Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(1, 1).FormulaR1C1 = "DuplicateID"
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(1, 1).FormulaR1C1 = "Duplicates"
        
        ' Update the extents
        FLastColumn = FLastColumn + 2
        
        iKPCol = FindColumn("KP")
        iEventCodeCol = FindColumn("Event Code")
        
        ' Copy the formula across the first column
        For i = 2 To FLastRow
            ' Rounding to 3dp to increase the chance of a collision.  I'd rather have more false positives here than issues in Coabis
            Cells(i, 2).Value = Format(Cells(i, iKPCol).Value, "0.0000") & "." & Cells(i, iEventCodeCol).Value
            Cells(i, 1).Formula = "=COUNTIF(B2:B" & FLastRow & ", B" & i & ")<>1"
        Next i
        Columns("A:B").EntireColumn.AutoFit
        
        ' And let all the CountIf's be resolved at once
        Application.Calculation = xlCalculationAutomatic
    End If
    
    ' Add a column to report issues
    If FindColumn("Has Issue?") = -1 Then
        Columns("A:A").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(1, 1).FormulaR1C1 = "Has Issue?"
        Range("A2").Select
        ActiveCell.FormulaR1C1 = "No"
        Selection.Copy
        Range("A2", "A" & FLastRow).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Columns("A:A").EntireColumn.AutoFit
        
        ' Update the extents
        FLastColumn = FLastColumn + 1
    End If
    
    ' Ensure AutoFilter enabled
    Range("A1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    ' Initialisation
    sLastSpan = ""
    sLastRockDump = ""
    sLastMattress = ""
    sLastTrench = ""
    sLastBurial = ""
    sLastInsp = ""
    
    ' Exactly where are the various columns?
    ' Note: If the column doesn't exist, FindColumn will return -1, which will
    '       cause an error to be raised later.  If any of these are -1 though,
    '       the error is actually here...
    iHeight = FindColumn("Height")
    iWidth = FindColumn("Width")
    iLength = FindColumn("Length")
    iContactCP = FindColumn("Contact CP")
    iCPCalibration = FindColumn("CP Calibration")
    iIncidentType = FindColumn("Incident Type Code")
    iIncident = FindColumn("Incident Code")
    iComment = FindColumn("Comment")
    iAnomaly = FindColumn("Anomaly")
    iAnomalyComment = FindColumn("Anomaly Comment")
    iLocation = FindColumn("Location")
    iClockPosition = FindColumn("Clock Position")
    iDistanceOff = FindColumn("Distance off")
    iDuplicatesCol = FindColumn("Duplicates")
    iKPCol = FindColumn("KP")
    iEventCodeCol = FindColumn("Event Code")
    
    ' Cycle over all the rows, looking for errors
    For i = 2 To FLastRow
        sIncidentType = Cells(i, iIncidentType).Value
        sIncident = Cells(i, iIncident).Value
        sLocation = Cells(i, iLocation).Value
        sComment = Trim(Cells(i, iComment).Value)
        sAnomalyComment = Trim(Cells(i, iAnomalyComment).Value)
        dCurrKP = Cells(i, iKPCol).Value
        
        If Cells(i, iDuplicatesCol).Value Then
            Cells(i, iDuplicatesCol).Select
            MarkSelectedAsIssue
        End If
        
        ' Event.Code Sameness?
        For iDupCheck = 0 To UBound(FDuplicateEventCodeChecks) - 1
            With FDuplicateEventCodeChecks(iDupCheck)
                If Cells(i, iEventCodeCol).Value = .Code Then
                    If (.lastRow <> -1) And (Abs(.LastKP - dCurrKP) <= FKPThresholdForSameness) Then
                        ' Mark both this and the previous FJ as an issue
                        Cells(i, iKPCol).Select
                        MarkSelectedAsIssue
                        
                        Cells(.lastRow, iKPCol).Select
                        MarkSelectedAsIssue
                    End If
                    
                    .lastRow = i
                    .LastKP = dCurrKP
                End If
            End With
        Next iDupCheck
        
        ' Incident Type Sameness?
        For iDupCheck = 0 To UBound(FDuplicateIncidentTypeChecks) - 1
            With FDuplicateIncidentTypeChecks(iDupCheck)
                If sIncidentType = .Code Then
                    If (.lastRow <> -1) And (Abs(.LastKP - dCurrKP) <= FKPThresholdForSameness) Then
                        ' Mark both this and the previous FJ as an issue
                        Cells(i, iKPCol).Select
                        MarkSelectedAsIssue
                        
                        Cells(.lastRow, iKPCol).Select
                        MarkSelectedAsIssue
                    End If
                    
                    .lastRow = i
                    .LastKP = dCurrKP
                End If
            End With
        Next iDupCheck
        
        ' CP Checks
        If (Cells(i, iContactCP).Value <> "0") Then
            i1 = Val(Cells(i, iContactCP).Value)
            
            ' Missing CP Calibration?
            If (Cells(i, iCPCalibration).Value = "0") Then
                Cells(i, iCPCalibration).Select
                MarkSelectedAsIssue
            End If
            
            ' Positive CP?
            If i1 > 0 Then
                Cells(i, iContactCP).Select
                MarkSelectedAsIssue
            End If
            
            ' Anomalous, but not marked?
            If (i1 < -1100) And (Cells(i, iAnomaly).Value <> "Yes") Then
                Cells(i, iAnomaly).Select
                MarkSelectedAsIssue
                
                Cells(i, iContactCP).Select
                MarkSelectedAsIssue
            End If
        End If
        
        ' Positive Cal?
        If Val(Cells(i, iCPCalibration).Value) > 0 Then
            Cells(i, iCPCalibration).Select
            MarkSelectedAsIssue
        End If
        
        ' Missing Span Height?
        ' If sIncidentType = "SP" Then
        '     If Cells(i, iHeight).Value = "0" Then
        '         Cells(i, iHeight).Select
        '         MarkSelectedAsIssue
        '     End If
        ' End If
        
        ' DB/DM Checks
        If (sIncidentType = "DB") Or (sIncidentType = "DM") Then
            ' Are Dimensions populated?
            iDimCount = 0
            
            If Cells(i, iLength).Value <> "0" Then
                iDimCount = iDimCount + 1
            End If
            
            If Cells(i, iWidth).Value <> "0" Then
                iDimCount = iDimCount + 1
            End If
            
            If Cells(i, iHeight).Value <> "0" Then
                iDimCount = iDimCount + 1
            End If
            
            If iDimCount < 2 Then
                If Cells(i, iLength).Value = "0" Then
                    Cells(i, iLength).Select
                    MarkSelectedAsIssue
                End If
                
                If Cells(i, iWidth).Value = "0" Then
                    Cells(i, iWidth).Select
                    MarkSelectedAsIssue
                End If
                
                If Cells(i, iHeight).Value = "0" Then
                    Cells(i, iHeight).Select
                    MarkSelectedAsIssue
                End If
            End If
            
            ' Does the clock position need to be populated?
            If (InStr(sLocation, "Touching") > 0) Or (InStr(sLocation, "top") > 0) Or (InStr(sLocation, "Under") > 0) Or (sLocation = "0 - Not Applicable") Then
                ' Has the clock position been populated?
                If Cells(i, iClockPosition).Value = "N/A" Then
                    Cells(i, iClockPosition).Select
                    MarkSelectedAsIssue
                End If
            End If
            
            ' Has the Location been populated?
            If (sIncidentType = "DB") And (sLocation = "0 - Not Applicable") Then
                Cells(i, iLocation).Select
                MarkSelectedAsIssue
            End If
            
            ' Is there a comment?
            If (sComment = "-") Or (sComment = "") Then
                Cells(i, iComment).Select
                MarkSelectedAsIssue
            End If
        End If
        
        ' If the clock position screwed?
        If Val(Cells(i, iClockPosition).Value) > 40000 Then
            Cells(i, iClockPosition).Select
            MarkSelectedAsIssue
        End If
        
        ' Is Distance Off required?
        If (sLocation = "1 - Right (Off pipe)") Or (sLocation = "5 - Left (Off pipe)") Then
            If Cells(i, iDistanceOff).Value = "0" Then
                Cells(i, iDistanceOff).Select
                MarkSelectedAsIssue
            End If
        End If
        
        ' Just a single 0?
        If (sComment = "0") Then
            Cells(i, iComment).Select
            MarkSelectedAsIssue
        End If
        
        ' Just a single 0?
        If (sAnomalyComment = "0") Then
            Cells(i, iAnomalyComment).Select
            MarkSelectedAsIssue
        End If
        
        ' Any wierd unicode?
        If (Not Text_IsLatin(sComment)) Then
            Cells(i, iComment).Select
            MarkSelectedAsIssue
        End If
        
        ' Any wierd unicode?
        If Not Text_IsLatin(sAnomalyComment) Then
            Cells(i, iAnomalyComment).Select
            MarkSelectedAsIssue
        End If
        
        ' If Anomaly, do we have Anomaly Text?
        If (Cells(i, iAnomaly).Value = "Yes") Then
            If (sAnomalyComment = "") Or (sAnomalyComment = "-") Then
                Cells(i, iAnomalyComment).Select
                MarkSelectedAsIssue
            End If
        End If
        
        ' Missing KP Length?
        ' Lazy Check - does the Incident end with an E or S?
        If (Right(sIncident, 1) = "S") Or (Right(sIncident, 1) = "E") Then
            ' Rats , there 's single codes that meet that requirement....
            ' Better exclude them....
            If (sIncident <> "OTS") And (sIncident <> "CRS") And (sIncident <> "BRS") Then
                i1 = FindColumn("KP Length")
                
                If Cells(i, i1).Value = "0" Then
                    Cells(i, i1).Select
                    MarkSelectedAsIssue
                End If
            End If
        End If
        
        ' Ensure Start and End is matched...
        If (sIncidentType = "SP") Then
            If sIncident = sLastSpan Then
                Cells(iLastSpan, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(iLastSpan, iIncident).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncident).Select
                MarkSelectedAsIssue
            End If
            
            sLastSpan = sIncident
            iLastSpan = i
        End If
        
        ' Ensure Start and End is matched...
        ' CVX confirm we will need to move Start Inspection so KP matches end Inspection...
        If False Then
            ' Arg.  processing this by hand is too painful.  My vote - hand hack the Coabis Export....
            If (sIncidentType = "IN") Then
                If sIncident = sLastInsp Then
                    Cells(iLastInsp, iIncidentType).Select
                    MarkSelectedAsIssue
                    
                    Cells(iLastInsp, iIncident).Select
                    MarkSelectedAsIssue
                    
                    Cells(i, iIncidentType).Select
                    MarkSelectedAsIssue
                    
                    Cells(i, iIncident).Select
                    MarkSelectedAsIssue
                End If
                
                sLastInsp = sIncident
                iLastInsp = i
            End If
        End If
        
        ' Ensure Start and End is matched...
        If (sIncidentType = "BU") Then
            If sIncident = sLastBurial Then
                Cells(iLastBurial, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(iLastBurial, iIncident).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncident).Select
                MarkSelectedAsIssue
            End If
            
            sLastBurial = sIncident
            iLastBurial = i
        End If
        
        ' Ensure Start and End is matched...
        If (sIncidentType = "ST") And (Left(sIncident, 2) = "RD") Then
            If sIncident = sLastRockDump Then
                Cells(iLastRockDump, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(iLastRockDump, iIncident).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncident).Select
                MarkSelectedAsIssue
            End If
            
            sLastRockDump = sIncident
            iLastRockDump = i
        End If
        
        ' Ensure Start and End is matched...
        If (sIncidentType = "ST") And (Left(sIncident, 2) = "PM") Then
            If sIncident = sLastMattress Then
                Cells(iLastMattress, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(iLastMattress, iIncident).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncident).Select
                MarkSelectedAsIssue
            End If
            
            sLastMattress = sIncident
            iLastMattress = i
        End If
    
        ' Ensure Start and End is matched...
        If (sIncidentType = "ST") And (Left(sIncident, 2) = "TR") Then
            If sIncident = sLastTrench Then
                Cells(iLastTrench, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(iLastTrench, iIncident).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncidentType).Select
                MarkSelectedAsIssue
                
                Cells(i, iIncident).Select
                MarkSelectedAsIssue
            End If
            
            sLastTrench = sIncident
            iLastTrench = i
        End If
    Next i
    
    ' We've added some columns.  Ensure the Header Formatting is correct
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    ' Finish by selecting a sane default position
    Range("A2").Select
End Sub

Sub ColorSelected()
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

Sub MarkSelectedAsIssue()
    ColorSelected
    
    Cells(ActiveCell.Row, 1).Value = "Yes"
    Cells(ActiveCell.Row, 1).Select
    
    ColorSelected
End Sub

Sub FormatActiveSheet()
    Dim i As Long
    
    ' Use the Basic Tidy as a base...
    Call BasicTidy(ActiveSheet)
    
    ' Format specific columns
    ForceFindExtents
    
    Columns(FindColumn("KP")).Select
    Selection.NumberFormat = "0.0000"
    
    Columns(FindColumn("Easting")).Select
    Selection.NumberFormat = "0.0"
    
    Columns(FindColumn("Northing")).Select
    Selection.NumberFormat = "0.0"
    
    Columns(FindColumn("Depth")).Select
    Selection.NumberFormat = "0.0"
    
    i = FindColumn("VWTimestamp")
    If i <> -1 Then
        Columns(i).Select
        Selection.NumberFormat = "0"
    End If
    
    Range(Columns(FindColumn("Temperature")), Columns(FindColumn("DCC"))).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range(Columns(FindColumn("Contact CP")), Columns(FindColumn("ROV Heading"))).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Columns(FindColumn("Comment")).Select
    Selection.WrapText = True
    
    Columns(FindColumn("Anomaly Comment")).Select
    Selection.WrapText = True

    Columns(FindColumn("Comment")).ColumnWidth = 86.43
    Columns(FindColumn("Anomaly Comment")).ColumnWidth = 52.86
    Columns(FindColumn("Incident Type Code")).ColumnWidth = 4.86
    Columns(FindColumn("Incident Code")).ColumnWidth = 7
    Columns(FindColumn("Temperature")).ColumnWidth = 7.71
    Columns(FindColumn("Altitude")).ColumnWidth = 5.14
    Range(Columns(FindColumn("Distance off")), Columns(FindColumn("Clock Position"))).Select
    Selection.HorizontalAlignment = xlCenter
    
    Cells.Select
    Cells.EntireRow.AutoFit
    
    Range("A2").Select
End Sub

Sub SetCommentsForBurialEvents()
    Dim i As Long
    Dim bInRockDump As Boolean
    Dim iIncidentCol, iCommentCol As Long
    Dim sIncident, sComment As String
   
    ' Set the Header row
    ForceFindExtents
    
    iIncidentCol = FindColumn("Incident")
    iCommentCol = FindColumn("Comment")
    bInRockDump = False
    
    For i = 2 To FLastRow
        Cells(i, 1).Select
        
        sIncident = Trim(Cells(i, iIncidentCol).Value)
        sComment = Trim(Cells(i, iCommentCol).Value)
        
        If sIncident = "Rock Dump Start" Then
            bInRockDump = True
        ElseIf sIncident = "Rock Dump End" Then
            bInRockDump = False
        End If
        
        If sIncident = "Burial Start" Then
            If (sComment = "-") Or (sComment = "") Then
                If bInRockDump Then
                    Cells(i, iCommentCol).Value = "Buried under rock dump"
                Else
                    Cells(i, iCommentCol).Value = "Buried under seabed"
                End If
            End If
        End If
    Next i
    
    ' Finish by selecting a sane default position
    Range("A2").Select
End Sub

Sub InterpolateColumn(iCol As Long, iStartRow As Long, iEndRow As Long, iCurrRow As Long, dPercent As Double)
    Dim d1 As Double, d2 As Double
    
    Cells(iCurrRow, iCol).Select
    
    d1 = Cells(iStartRow, iCol).Value
    d2 = Cells(iEndRow, iCol).Value
    Cells(iCurrRow, iCol).Value = d1 + (dPercent * (d2 - d1))
    
    ColorSelected
End Sub

Sub InterpolatePositionFromMBES()
    Dim iDateCol As Long
    Dim iTimeCol As Long
    Dim iIncidentCol As Long, iFixedCol As Long
    Dim iEastingCol As Long
    Dim iNorthingCol As Long
    Dim iDepthCol As Long
    Dim iDCCCol As Long
    Dim iKPCol As Long
    Dim iMBESKPCol As Long
    Dim iRow As Long, iCurrRow As Long, iStartRow As Long, iEndRow As Long
    Dim dPercent As Double
    Dim sIncidentCode As String, sFixed As String
    Dim dtStart As Date, dtEnd As Date, dtCurr As Date
    
    ' Set the Header row
    ForceFindExtents
    
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    iIncidentCol = FindColumn("Incident Code")
    iEastingCol = FindColumn("Easting")
    iNorthingCol = FindColumn("Northing")
    iDepthCol = FindColumn("Depth")
    iDCCCol = FindColumn("DCC")
    iKPCol = FindColumn("KP")
    iMBESKPCol = FindColumn("MBES_KP")
    iFixedCol = FindColumn("Fixed")
    
    iStartRow = -1
    iEndRow = -1
    
    For iRow = 2 To FLastRow
    'For iRow = 3053 To 3060
        sFixed = UCase(Trim(Cells(iRow, iFixedCol).Value))
        
        If sFixed = "Y" Then
            Cells(iRow, iIncidentCol).Select
            
            If iStartRow <> -1 Then
                iEndRow = iRow
            End If
                
            If iEndRow <> -1 Then
                dtStart = Cells(iStartRow, iDateCol).Value + Cells(iStartRow, iTimeCol).Value
                dtEnd = Cells(iEndRow, iDateCol).Value + Cells(iEndRow, iTimeCol).Value
                
                For iCurrRow = (iStartRow + 1) To (iEndRow - 1)
                    dtCurr = Cells(iCurrRow, iDateCol).Value + Cells(iCurrRow, iTimeCol).Value
                    dPercent = (dtCurr - dtStart) / (dtEnd - dtStart)
                    
                    Call InterpolateColumn(iMBESKPCol, iStartRow, iEndRow, iCurrRow, dPercent)
                    Call InterpolateColumn(iEastingCol, iStartRow, iEndRow, iCurrRow, dPercent)
                    Call InterpolateColumn(iNorthingCol, iStartRow, iEndRow, iCurrRow, dPercent)
                    Call InterpolateColumn(iDepthCol, iStartRow, iEndRow, iCurrRow, dPercent)
                    Call InterpolateColumn(iDCCCol, iStartRow, iEndRow, iCurrRow, dPercent)
                Next iCurrRow
            End If
            
            iStartRow = iRow
        End If
    Next iRow
End Sub

Sub SetInspectionEndPosToNextInspectionStart()
    Dim iIncidentCol As Long
    Dim iKPCol As Long, iDateCol As Long, iTimeCol As Long, iContCP As Long
    Dim iRow As Long
    Dim sIncidentCode As String
    Dim iLastEndRow As Long
    Dim sLastCode As String
    Dim bDirection As Boolean
    Dim sCodeTo As String
    Dim iColumn As Long
    Dim dOffset As Double
    
    
    bDirection = True ' Ascending
    ' bDirection = False ' Descending
    
    If bDirection Then
        sCodeTo = "INE"
        dOffset = -0.0001
    Else
        sCodeTo = "INS"
        dOffset = 0.0001
    End If
    
    ' Set the Header row
    ForceFindExtents
    
    iIncidentCol = FindColumn("Incident Code")
    iKPCol = FindColumn("KP")
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    iContCP = FindColumn("Continuous CP")
    
    '   Data must be sorted by time first
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iDateCol), Cells(FLastRow, iDateCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iTimeCol), Cells(FLastRow, iTimeCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
    iLastEndRow = -1
    sLastCode = "XXX"
    
    For iRow = 2 To FLastRow - 1
        sIncidentCode = UCase(Trim(Cells(iRow, iIncidentCol).Value))
        
        If (sIncidentCode = "INE") Or (sIncidentCode = "INS") Then
            Cells(iRow, iKPCol).Select
            
            If sIncidentCode = sLastCode Then
                MsgBox ("Double IN code")
                Exit For
            End If
            
            If (sIncidentCode = sCodeTo) Then
                iLastEndRow = iRow
            ElseIf (iLastEndRow <> -1) Then
                If (Abs(Cells(iLastEndRow, iKPCol).Value - Cells(iRow, iKPCol).Value) > 0.0002) Then
                    Cells(iLastEndRow, iKPCol).Value = Cells(iRow, iKPCol).Value + dOffset ' This start slightly offset of previous end to ensure Coabis doesn't die
                    Cells(iLastEndRow, iKPCol).Select
                    ColorSelected
                    
                    ' iColumn = iDateCol: GoSub CopyCell
                    ' iColumn = iTimeCol: GoSub CopyCell
                    iColumn = iKPCol + 1: GoSub CopyCell ' Easting
                    iColumn = iKPCol + 2: GoSub CopyCell ' Northing
                    iColumn = iKPCol + 3: GoSub CopyCell ' Depth
                    iColumn = iKPCol + 4: GoSub CopyCell ' Temp
                    iColumn = iKPCol + 5: GoSub CopyCell ' Alt
                    iColumn = iKPCol + 6: GoSub CopyCell ' DCC
                    iColumn = iContCP: GoSub CopyCell     ' iContCP
                    iColumn = iContCP + 1: GoSub CopyCell ' FG
                    iColumn = iContCP + 2: GoSub CopyCell ' DOB
                    iColumn = iContCP + 3: GoSub CopyCell ' Pitch
                    iColumn = iContCP + 4: GoSub CopyCell ' Roll
                    iColumn = iContCP + 5: GoSub CopyCell ' Heading
                End If
            End If
            
            sLastCode = sIncidentCode
        End If
    Next iRow
    
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iKPCol), Cells(FLastRow, iKPCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
    Exit Sub
    
CopyCell:
    Cells(iLastEndRow, iColumn).Value = Cells(iRow, iColumn).Value
    Cells(iLastEndRow, iColumn).Select
    ColorSelected
    
    Return
End Sub

Sub FixInspectionEndTime()
    Dim iRow As Long
    Dim iDateCol As Long, iTimeCol As Long
    
     ' Set the Header row
    ForceFindExtents
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    
    For iRow = 2 To FLastRow - 1
        If Cells(iRow, 32) = "IN.INE" Then
            Cells(iRow, iDateCol).Value = Cells(iRow - 1, iDateCol).Value
            Cells(iRow, iTimeCol).Value = DateAdd("s", 1, Cells(iRow - 1, iTimeCol).Value)
            Cells(iRow, iTimeCol).Select
            ColorSelected
        End If
    Next iRow
End Sub

Sub InterpolateTimeFromKP()
    Dim iDateCol As Long
    Dim iTimeCol As Long
    Dim iIncidentCol As Long, iFixedCol As Long
    Dim iDCCCol As Long
    Dim iKPCol As Long
    Dim iRow As Long, iCurrRow As Long, iStartRow As Long, iEndRow As Long
    Dim dPercent As Double
    Dim sIncidentCode As String, sFixed As String
    Dim dtStart As Date, dtEnd As Date, dtCurr As Date
    Dim dStart As Double, dEnd As Double, dCurr As Double
    Dim DCCStart As Double, DCCEnd As Double, DCCCurr As Double
    
    ' Set the Header row
    ForceFindExtents
    
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    iIncidentCol = FindColumn("Incident Code")
    iDCCCol = FindColumn("DCC")
    iKPCol = FindColumn("KP")
    iFixedCol = FindColumn("Fixed")
    
    iStartRow = -1
    iEndRow = -1
    
    For iRow = 2 To FLastRow
        sFixed = UCase(Trim(Cells(iRow, iFixedCol).Value))
        
        If sFixed <> "NEW" Then
            Cells(iRow, iIncidentCol).Select
            
            If iStartRow <> -1 Then
                iEndRow = iRow
            End If
                
            If iEndRow <> -1 Then
                dStart = Cells(iStartRow, iKPCol).Value
                dEnd = Cells(iEndRow, iKPCol).Value
                
                DCCStart = Cells(iStartRow, iDCCCol).Value
                DCCEnd = Cells(iEndRow, iDCCCol).Value
                
                dtStart = Cells(iStartRow, iDateCol).Value + Cells(iStartRow, iTimeCol).Value
                dtEnd = Cells(iEndRow, iDateCol).Value + Cells(iEndRow, iTimeCol).Value
                
                For iCurrRow = (iStartRow + 1) To (iEndRow - 1)
                    dCurr = Cells(iCurrRow, iKPCol).Value
                    dPercent = (dCurr - dStart) / (dEnd - dStart)
                    
                    dtCurr = dtStart + ((dtEnd - dtStart) * dPercent)
                    
                    Cells(iCurrRow, iDateCol).Value = DateValue(dtCurr)
                    Cells(iCurrRow, iTimeCol).Value = TimeValue(dtCurr)
                    
                    Cells(iCurrRow, iTimeCol).Select
                    ColorSelected
                    
                    Cells(iCurrRow, iDateCol).Select
                    ColorSelected
                    
                    If Cells(iCurrRow, iDCCCol).Value = "" Then
                        DCCCurr = DCCStart + ((DCCEnd - DCCStart) * dPercent)
                        
                        Cells(iCurrRow, iDCCCol).Value = DCCCurr
                        Cells(iCurrRow, iDCCCol).Select
                        ColorSelected
                    End If
                Next iCurrRow
            End If
            
            iStartRow = iRow
        End If
    Next iRow
End Sub

Sub ApplyDM_WTC_Hack()
    Dim iEventCodeCol As Long
    Dim iHeight, iWidth, iLength As Long
    Dim iClockPosition As Long
    Dim iComment As Long
    
    Dim sClockPosition As String, sComment As String
    Dim sStart As String, sEnd As String
    
    Dim dOneClockLength As Double, dClockLength As Double
    Dim dStart As Double, dEnd As Double
    Dim i As Long
    
    
    ' One clock position on a 1m diameter pipeline is approx 250mm
    dOneClockLength = 0.25
    
    ' Set the Header row
    ForceFindExtents
    
    ' Find the columns we'll be using
    iHeight = FindColumn("Height")
    iWidth = FindColumn("Width")
    iLength = FindColumn("Length")
    iEventCodeCol = FindColumn("Event Code")
    iComment = FindColumn("Comment")
    iClockPosition = FindColumn("Clock Position")
    
    
    For i = 2 To FLastRow
        If Cells(i, iEventCodeCol).Value = "DM.WTC" Then
            ' Only apply this hack under extreme conditions
            If (Cells(i, iHeight).Value = "0") And (Cells(i, iWidth).Value = "0") And (Cells(i, iLength).Value = "0") And (Cells(i, iHeight).Value = "0") Then
                sComment = Trim(LCase(Cells(i, iComment).Value))
                sClockPosition = Trim(LCase(Cells(i, iClockPosition).Value))
                
                If ((sComment = "-") Or (sComment = "")) And (Trim(Cells(i, iClockPosition).Value <> "")) Then
                    sStart = Trim(Text_Between(sClockPosition, "", " to"))
                    sEnd = Trim(Text_Between(sClockPosition, "to", ""))
                    
                    If InStr(" ", sEnd) Then
                        sEnd = Trim(Text_Between(sEnd, "", " "))
                    End If
                    
                    dStart = Val(sStart)
                    dEnd = Val(sEnd)
                    
                    If dStart > 6 And dEnd > 6 Then
                        dClockLength = dEnd - dStart
                    ElseIf dStart > 6 And dEnd < 6 Then
                        dClockLength = (12 - sStart) + dEnd
                    ElseIf dStart < 6 And dEnd > 6 Then
                        dClockLength = dEnd - dStart
                    ElseIf dStart < 6 And dEnd < 6 Then
                        dClockLength = dEnd - dStart
                    End If
                    
                    ' Now populate the hack values
                    Cells(i, iHeight).Value = "0.1"
                    Cells(i, iLength).Value = "0.1"
                    Cells(i, iWidth).Value = Round(dClockLength, 1) * dOneClockLength
                    Cells(i, iComment).Value = "Minor weightcoat damage at Fieldjoint along edge"

                    Cells(i, iComment).Select
                    ColorSelected
                End If
            End If
        End If
    Next i

End Sub

Sub UpdateKPLength()
    Dim iIncidentCol As Long
    Dim iKPCol As Long, iKPLengthCol As Long
    Dim iRow As Long, iCode As Long
    Dim iDateCol As Long, iTimeCol As Long
    Dim sIncidentCode As String
    Dim dKPLength As Double
    
    ' Set the Header row
    ForceFindExtents
    
    iIncidentCol = FindColumn("Incident Code")
    iKPCol = FindColumn("KP")
    iKPLengthCol = FindColumn("KP Length")
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    
    ReDim FIncidentPairs(1)
    
    If FAscendingInspection Then
        FIncidentPairs(0).StartCode = "INS"
        FIncidentPairs(0).EndCode = "INE"
        FIncidentPairs(0).StartKP = -1
    Else
        FIncidentPairs(0).StartCode = "INE"
        FIncidentPairs(0).EndCode = "INS"
        FIncidentPairs(0).StartKP = -1
    End If
    
    ' Process the time based pairs first
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iDateCol), Cells(FLastRow, iDateCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iTimeCol), Cells(FLastRow, iTimeCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    For iRow = 2 To FLastRow
        sIncidentCode = UCase(Trim(Cells(iRow, iIncidentCol).Value))
        
        For iCode = 0 To UBound(FIncidentPairs) - 1
            If sIncidentCode = FIncidentPairs(iCode).StartCode Then
                FIncidentPairs(iCode).StartKP = Cells(iRow, iKPCol).Value
                FIncidentPairs(iCode).StartRow = iRow
            ElseIf sIncidentCode = FIncidentPairs(iCode).EndCode Then
                If FIncidentPairs(iCode).StartKP = -1 Then
                    MsgBox ("Missing Start Code for " & FIncidentPairs(iCode).EndCode & " on Row " & iRow)
                    Exit For
                End If
                
                dKPLength = Abs(Round(1000 * (Cells(iRow, iKPCol).Value - FIncidentPairs(iCode).StartKP), 1))
                
                If Abs(Cells(iRow, iKPLengthCol).Value - dKPLength) >= 0.05 Then
                    Cells(iRow, iKPLengthCol).Value = dKPLength
                
                    Cells(iRow, iKPLengthCol).Select
                    ColorSelected
                End If
                
                If Abs(Cells(FIncidentPairs(iCode).StartRow, iKPLengthCol).Value - dKPLength) >= 0.05 Then
                    Cells(FIncidentPairs(iCode).StartRow, iKPLengthCol).Value = dKPLength
                    
                    Cells(FIncidentPairs(iCode).StartRow, iKPLengthCol).Select
                    ColorSelected
                End If
                
                FIncidentPairs(iCode).StartKP = -1
            End If
        Next iCode
    Next iRow
    
    ' Process the KP based pairs
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iKPCol), Cells(FLastRow, iKPCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    ReDim FIncidentPairs(6)
    
'    If FAscendingInspection Then
        FIncidentPairs(0).StartCode = "BUS"
        FIncidentPairs(0).EndCode = "BUE"
        FIncidentPairs(0).StartKP = -1
        
        FIncidentPairs(1).StartCode = "RDS"
        FIncidentPairs(1).EndCode = "RDE"
        FIncidentPairs(1).StartKP = -1
        
        FIncidentPairs(2).StartCode = "SPS"
        FIncidentPairs(2).EndCode = "SPE"
        FIncidentPairs(2).StartKP = -1
        
        FIncidentPairs(3).StartCode = "WCS"
        FIncidentPairs(3).EndCode = "WCE"
        FIncidentPairs(3).StartKP = -1
    
        FIncidentPairs(4).StartCode = "TRS"
        FIncidentPairs(4).EndCode = "TRE"
        FIncidentPairs(4).StartKP = -1
        
        FIncidentPairs(5).StartCode = "PMS"
        FIncidentPairs(5).EndCode = "PME"
        FIncidentPairs(5).StartKP = -1
'    Else
'        FIncidentPairs(0).StartCode = "BUE"
'        FIncidentPairs(0).EndCode = "BUS"
'        FIncidentPairs(0).StartKP = -1
'
'        FIncidentPairs(1).StartCode = "RDE"
'        FIncidentPairs(1).EndCode = "RDS"
'        FIncidentPairs(1).StartKP = -1
'
'        FIncidentPairs(2).StartCode = "SPE"
'        FIncidentPairs(2).EndCode = "SPS"
'        FIncidentPairs(2).StartKP = -1
'
'        FIncidentPairs(3).StartCode = "WCE"
'        FIncidentPairs(3).EndCode = "WCS"
'        FIncidentPairs(3).StartKP = -1
'
'        FIncidentPairs(4).StartCode = "TRE"
'        FIncidentPairs(4).EndCode = "TRS"
'        FIncidentPairs(4).StartKP = -1
'
'        FIncidentPairs(5).StartCode = "PME"
'        FIncidentPairs(5).EndCode = "PMS"
'        FIncidentPairs(5).StartKP = -1
'    End If
    
    For iRow = 2 To FLastRow
        sIncidentCode = UCase(Trim(Cells(iRow, iIncidentCol).Value))
        
        For iCode = 0 To UBound(FIncidentPairs) - 1
            If sIncidentCode = FIncidentPairs(iCode).StartCode Then
                FIncidentPairs(iCode).StartKP = Cells(iRow, iKPCol).Value
                FIncidentPairs(iCode).StartRow = iRow
            ElseIf sIncidentCode = FIncidentPairs(iCode).EndCode Then
                If FIncidentPairs(iCode).StartKP = -1 Then
                    MsgBox ("Missing Start Code for " & FIncidentPairs(iCode).EndCode & " on Row " & iRow)
                    Exit For
                End If
                
                dKPLength = Abs(Round(1000 * (Cells(iRow, iKPCol).Value - FIncidentPairs(iCode).StartKP), 1))
                
                If Abs(Cells(iRow, iKPLengthCol).Value - dKPLength) >= 0.05 Then
                    Cells(iRow, iKPLengthCol).Value = dKPLength
                
                    Cells(iRow, iKPLengthCol).Select
                    ColorSelected
                End If
                
                If Abs(Cells(FIncidentPairs(iCode).StartRow, iKPLengthCol).Value - dKPLength) >= 0.05 Then
                    Cells(FIncidentPairs(iCode).StartRow, iKPLengthCol).Value = dKPLength
                    
                    Cells(FIncidentPairs(iCode).StartRow, iKPLengthCol).Select
                    ColorSelected
                End If
                
                FIncidentPairs(iCode).StartKP = -1
            End If
        Next iCode
    Next iRow
End Sub

Sub CalculateNumberOfRBPerSpan()
    Dim iIncidentCol As Long, iRBCountCol As Long
    Dim iRow As Long, iStartRow As Long, iRBCount As Long
    Dim sIncidentCode As String
    
    ' Set the Header row
    ForceFindExtents
    
    iIncidentCol = FindColumn("Incident Code")
    iRBCountCol = FindColumn("RB Count")
    
    iStartRow = -1
    iRBCount = 0
    
    For iRow = 2 To FLastRow
        sIncidentCode = UCase(Trim(Cells(iRow, iIncidentCol).Value))
        
        If (sIncidentCode = "SPS") Or (sIncidentCode = "SPE") Then
            Cells(iRow, iIncidentCol).Select
            
            If (sIncidentCode = "SPS") Then
                If (iStartRow <> -1) Then
                    Err.Raise 555, "CalculateNumberOfRBPerSpan", "Mismatched Start Span"
                End If
                
                iRBCount = 0
                iStartRow = iRow
            Else
                If (iStartRow = -1) Then
                    Err.Raise 555, "CalculateNumberOfRBPerSpan", "Mismatched Start End"
                End If
                
                Cells(iRow, iRBCountCol).Value = iRBCount
                Cells(iStartRow, iRBCountCol).Value = iRBCount
                
                iRBCount = 0
                iStartRow = -1
            End If
        ElseIf (sIncidentCode = "RKB") Then
            If iStartRow <> -1 Then
                iRBCount = iRBCount + 1
            End If
        End If
    Next iRow
End Sub

Public Sub ProcessVWCoabisExport()
    Dim i As Long
    Dim dTemp As Date
    Dim sTemp As String
    
    Call BasicTidy(ActiveSheet)
    
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    Rows("1:1").Select
    ForceFindExtents
    
    For i = 2 To FLastRow
        Cells(i, 5).Select
        sTemp = Trim(Cells(i, 5).Value)
        dTemp = DateValue(Left(sTemp, 10))
        dTemp = dTemp + TimeValue(Right(sTemp, 8))
        Cells(i, 5).Value = dTemp
        
        Cells(i, 32).Select
        sTemp = Trim(Cells(i, 32).Value)
        dTemp = DateValue(Left(sTemp, 10))
        Cells(i, 32).Value = dTemp
    Next i
    
    Columns(5).Select
    Selection.NumberFormat = "dd/mm/yyyy HH:mm:ss"
    
    Columns(32).Select
    Selection.NumberFormat = "yyyy/mm/dd"
    
    ActiveCell.SpecialCells(xlLastCell).Select
    For i = ActiveCell.Row To FLastRow + 1 Step -1
        Rows(i).Select
        Rows(i).Delete
    Next i
    
    Cells(2, 1).Select
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs filename:=Text_Replace(ActiveWorkbookLocalFilename, ".xlsx", ".xls"), FileFormat:=xlExcel8, Local:=True
    ActiveWorkbook.CheckCompatibility = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
End Sub

Sub InterpolateMiddleRow()
    ' This code is designed to work on the Coabis Pipeline Import spreadsheet
    
    Dim iStartRow As Long, iEndRow As Long, iNewRow As Long
    Dim iCol As Long
    Dim iKPCol As Long, iTimeCol As Long, iTypeCol As Long, iCodeCol As Long, iCommentCol As Long, iClockCol As Long
    Dim dStartKP As Double, dEndKP As Double, dCurrKP As Double
    Dim vStart, vEnd As Variant   ' Deliberately left as Variants
    Dim dStart As Double, dEnd As Double
    
    If Selection.Rows.Count <> 3 Then
      MsgBox ("Please ensure you have selected the three rows you wish to interpolate (and only those three rows)." & vbCrLf & "The first and last row must contain the original data, and the middle row must contain KP")
      Exit Sub
    End If
    
    iStartRow = Selection.Rows(1).Row
    iNewRow = Selection.Rows(2).Row
    iEndRow = Selection.Rows(3).Row
    
    Rows(iStartRow).Select
    Selection.Interior.Pattern = xlNone
    Selection.Font.Color = -10526881
    
    Rows(iNewRow).Select
    Selection.Interior.Color = 10092543
    
    Rows(iEndRow).Select
    Selection.Interior.Pattern = xlNone
    Selection.Font.Color = -10526881
    
    ForceFindExtents
    
    iKPCol = FindColumn("KP")
    iTimeCol = FindColumn("Incident Date and Time")
    iTypeCol = FindColumn("Incident Type")
    iCodeCol = FindColumn("Incident Code")
    iCommentCol = FindColumn("Comment")
    iClockCol = FindColumn("Clock_Position")
    
    If Trim(Cells(iNewRow, iKPCol).Value) = "" Then
        MsgBox ("You need to populate the KP in the interpolated Row.")
        Exit Sub
    End If
        
    dStartKP = Cells(iStartRow, iKPCol).Value
    dCurrKP = Cells(iNewRow, iKPCol).Value
    dEndKP = Cells(iEndRow, iKPCol).Value
    
    Cells(iStartRow, 1).Select
    
    For iCol = 1 To FLastColumn
        Cells(iNewRow, iCol).Select
        
        vStart = Cells(iStartRow, iCol).Value
        vEnd = Cells(iEndRow, iCol).Value
        
        If iCol = iKPCol Then
          ' Do Nothing - this is our baseline Column
        
        ElseIf (iCol = iCommentCol) Then
            If (Trim(Cells(iNewRow, iCol).Value) = "") Then
                Cells(iNewRow, iCol).Value = "Not recorded during original inspection.  Insertion method: interpolation by KP"
            End If
        
        ElseIf iCol = iTypeCol Then
            Cells(iNewRow, iCol).Value = "FT"
        
        ElseIf iCol = iCodeCol Then
            Cells(iNewRow, iCol).Value = "CRS"
        
        ElseIf iCol = iClockCol Then
            Cells(iNewRow, iCol).Value = "N/A"
        
        ElseIf vStart = vEnd Then
            ' Only interpolate if we need to
            Cells(iNewRow, iCol).Value = vStart
        
        ElseIf iCol = iTimeCol Then
            dStart = vStart
            dEnd = vEnd
            Cells(iNewRow, iCol).Value = InterpolateByDouble(dStartKP, dEndKP, dCurrKP, dStart, dEnd)
        
        ElseIf TypeName(vStart) = "Double" Then
            dStart = vStart
            dEnd = vEnd
            Cells(iNewRow, iCol).Value = InterpolateByDouble(dStartKP, dEndKP, dCurrKP, dStart, dEnd)
        
        ElseIf TypeName(vStart) = "String" Then
            Cells(iNewRow, iCol).Value = "Error: Unable to Interpolate different text"
            ColorSelected
        
        Else
            Cells(iNewRow, iCol).Value = "Error: Data Type " & TypeName(vStart) & " not handled"
            ColorSelected
        End If
    Next iCol
    
    Cells(iNewRow, 1).Select
    Rows(iNewRow).Select
   
    MsgBox ("Finished." & vbCrLf & vbCrLf & "Please review entire new row.  Errors will be highlighted in, err, purple-ish." & vbCrLf & vbCrLf & "When finished please delete the original rows, leaving only the new row" & vbCrLf & vbCrLf & "We hope you enjoyed using this service.  Have a nice day")
End Sub
