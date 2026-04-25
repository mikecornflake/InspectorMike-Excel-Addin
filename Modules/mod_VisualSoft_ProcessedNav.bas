Attribute VB_Name = "mod_VisualSoft_ProcessedNav"
Option Explicit
Option Private Module

Public FBaseFolder As String
Public fCurrent As String
Public FTrack As String
Public FHistorical As String
Public FCode As String
Public FMatchCode As String

Private sInitialisationHack As String

Private Sub InitialiseGlobalVars()
'
' This can be called multiple times, it'll only do the main code once
'

    If sInitialisationHack = vbNullString Then
        ' Place all Initialisation code here
        FBaseFolder = "H:\Documents\14UI1 - Event Processing\"
        fCurrent = "01 - From VW (Offshore).csv"
        FTrack = "14UI1_kp-00.016_kp006.381_LASER.csv"
        FHistorical = "14UI1 - Anodes.xlsx"
        FCode = "AN"
        FMatchCode = "-1"
        
        sInitialisationHack = "Initialised"
    End If
End Sub

Public Sub VW_ShowMerge()
    InitialiseGlobalVars
    
    frmEventMerge.Show
End Sub

Public Sub VW_MergeEvents()
'
'  Very much prototype code
'
    Dim sBaseFolder As String, sHistoricalName As String, sNewName As String, sTrackName As String, sCurrentName As String
    
    Dim sIncidentTypeCode As String, sIncidentCode As String, iIncidentTypeCol As Long
    
    Dim iWorkingCol1 As Long, iWorkingCol2 As Long, iWorkingCol3 As Long
    Dim iWorkingCol4 As Long, iWorkingCol5 As Long, iWorkingCol6 As Long
    Dim iWorkingCol7 As Long, iWorkingCol8 As Long, iWorkingCol9 As Long
    Dim iWorkingCol10 As Long, iWorkingCol11 As Long, iWorkingCol12 As Long
    Dim iWorkingCol13 As Long
    
    Dim iHistoricalCol1 As Long, iHistoricalCol2 As Long, iHistoricalCol3 As Long
    
    Dim iKPCol As Long
    
    Dim iWorkingLastColumn As Long
    Dim iWorkingLastRow As Long, iEventsLastRow As Long, iTrackLastRow As Long
    
    Dim dThisKP As Double, dOtherKP As Double
    
    Dim i As Long
    Dim iRowsDeleted As Long
    
    Dim wbCurrent As Workbook, wbHistorical As Workbook
    Dim wsEvents As Worksheet, wsWorking As Worksheet, wsTrack As Worksheet, wsQC As Worksheet
    
    ' These change from processing job to processing job
    
    'sBaseFolder = "H:\Documents\Proc\6UJM - 01 Inputs\"
    'sCurrentName = sBaseFolder & "6UJM - VW.csv"
    'sTrackName = sBaseFolder & "6UJM_kp128.757_kp130.380_LASER.csv"
    'sHistoricalName = sBaseFolder & "6UJM - Field Joints.xlsx"
    'sIncidentTypeCode = "FJ"
    
    sBaseFolder = FBaseFolder
    sCurrentName = FBaseFolder & fCurrent
    sTrackName = FBaseFolder & FTrack
    sHistoricalName = FBaseFolder & FHistorical
    sIncidentTypeCode = FCode
    
    sNewName = Text_Replace(sCurrentName, ".csv", "- " & sIncidentTypeCode & " - " & Format(Date, "yyyy-mm-ss") & " " & Format(Time, "hhmmss") & ".xlsx")
    
    Workbooks.Open filename:=sCurrentName, Format:=xlDelimited, Local:=True
    ActiveWorkbook.SaveAs filename:=sNewName, FileFormat:=xlWorkbookDefault
    Set wbCurrent = ActiveWorkbook
    
    TidyVWExcelExport
    
    Set wsEvents = wbCurrent.ActiveSheet
    
    Set wsWorking = wbCurrent.Sheets.Add(After:=wsEvents)
    wsWorking.Name = sIncidentTypeCode
    
    Set wsTrack = wbCurrent.Sheets.Add(After:=wsWorking)
    wsTrack.Name = "Track"
    
    Set wsQC = wbCurrent.Sheets.Add(After:=wsTrack)
    wsQC.Name = "QC"
    
    wsEvents.Activate
    ForceFindExtents
    iIncidentTypeCol = FindColumn("Incident Type Code")
    
    
    ' Apply filter
    wsEvents.Range(Cells(1, 1), Cells(FLastRow, FLastColumn)).AutoFilter Field:=iIncidentTypeCol, Criteria1:=sIncidentTypeCode
    
    ' Copy filter to new tab
    wsEvents.Range(Cells(1, 1), Cells(FLastRow, FLastColumn)).Select
    Selection.Copy
    
    wsWorking.Activate
    ActiveSheet.Paste
    TidyVWExcelExport
    
    ForceFindExtents
    iWorkingCol1 = FindColumn("KP")
    iWorkingCol2 = FindColumn("Incident Type")
    iWorkingCol3 = FindColumn("Incident")
    iWorkingCol4 = FindColumn("Incident Type Code")
    iWorkingCol5 = FindColumn("Incident Code")
    iWorkingCol6 = FindColumn("Comment")
    iWorkingCol7 = FindColumn("Event Code")
    
    iWorkingCol8 = FindColumn("Location")
    iWorkingCol9 = FindColumn("Clock Position")
    iWorkingCol10 = FindColumn("Anomaly")
    iWorkingCol11 = FindColumn("Anomaly Comment")
    iWorkingCol12 = FindColumn("Contact CP")
    iWorkingCol13 = FindColumn("CP Calibration")
    
    iWorkingLastRow = FLastRow
    iWorkingLastColumn = FLastColumn
    
    ' Open Historical Events
    Set wbHistorical = Workbooks.Open(sHistoricalName)
    
    ForceFindExtents
    iHistoricalCol1 = FindColumn("KP")
    iHistoricalCol2 = FindColumn("INCIDENT TYPE")
    iHistoricalCol3 = FindColumn("INCIDENT")
    
    iEventsLastRow = FLastRow
    
    For i = 2 To iEventsLastRow
        sIncidentCode = Cells(i, iHistoricalCol3).Value
        
        iWorkingLastRow = iWorkingLastRow + 1
        wsWorking.Cells(iWorkingLastRow, iWorkingCol1).Value = Cells(i, iHistoricalCol1).Value
        wsWorking.Cells(iWorkingLastRow, iWorkingCol2).Value = IncidentType(sIncidentTypeCode)
        wsWorking.Cells(iWorkingLastRow, iWorkingCol3).Value = Incident(sIncidentTypeCode, sIncidentCode)
        wsWorking.Cells(iWorkingLastRow, iWorkingCol4).Value = sIncidentTypeCode
        wsWorking.Cells(iWorkingLastRow, iWorkingCol5).Value = sIncidentCode
        wsWorking.Cells(iWorkingLastRow, iWorkingCol6).Value = "Position obtained from As-Built results.  Not seen this campaign"
        wsWorking.Cells(iWorkingLastRow, iWorkingCol7).Value = sIncidentTypeCode & "." & sIncidentCode
        
        ' fill in the blanks that VW seems to need...
        wsWorking.Cells(iWorkingLastRow, iWorkingCol8).Value = "0 - Not Applicable"
        wsWorking.Cells(iWorkingLastRow, iWorkingCol9).Value = "N/A"
        wsWorking.Cells(iWorkingLastRow, iWorkingCol10).Value = "No"
        wsWorking.Cells(iWorkingLastRow, iWorkingCol11).Value = "-"
        wsWorking.Cells(iWorkingLastRow, iWorkingCol12).Value = "0"
        wsWorking.Cells(iWorkingLastRow, iWorkingCol13).Value = "0"
    Next i
    
    wsWorking.Activate
    
    ' Now sort by KP
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iWorkingCol1), Cells(iWorkingLastRow, iWorkingCol1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    iRowsDeleted = 0
    
    ' Now remove any new row that is within 3m of an existing entry
    For i = iWorkingLastRow To 3 Step -1
        Cells(i, 1).Select
        
        ' If is new
        If Trim(Cells(i, 1).Value) = "" Then
            dThisKP = Cells(i, iWorkingCol1).Value
            dOtherKP = Cells(i - 1, iWorkingCol1).Value
            
            If 1000 * Abs(dThisKP - dOtherKP) <= 3 Then
                Rows(i).Delete
                iRowsDeleted = iRowsDeleted + 1
            End If
        End If
    Next i
    
    ' Now remove any new row that is within 3m of an existing entry, looking the other way
    For i = iWorkingLastRow - 1 To 2 Step -1
        Cells(i, 1).Select
        
        ' If is new
        If Trim(Cells(i, 1).Value) = "" Then
            dThisKP = Cells(i, iWorkingCol1).Value
            dOtherKP = Cells(i + 1, iWorkingCol1).Value
            
            If 1000 * Abs(dThisKP - dOtherKP) <= 3 Then
                Rows(i).Delete
                iRowsDeleted = iRowsDeleted + 1
            End If
        End If
    Next i
    
    ' Finally, remove all the existing events.
    ' Now remove any new row that is within 3m of an existing entry, looking the other way
    For i = iWorkingLastRow To 2 Step -1
        Cells(i, 1).Select
        
        ' If NOT new
        If Trim(Cells(i, 1).Value) <> "" Then
            Rows(i).Delete
                iRowsDeleted = iRowsDeleted + 1
        End If
    Next i
    
    ForceFindExtents
    iWorkingLastRow = FLastRow
    
    wbHistorical.Close SaveChanges:=False
    wbCurrent.Save
    
    ' At this time, all that is left on the Working Tabsheet is a list of the event codes which will need to be inserted to the Current copy
    ' What is missing right now though, is Time.  For this, we will need the Reprocessed Track
    wsTrack.Activate
    
    ' http://stackoverflow.com/questions/12197274/is-there-a-way-to-import-data-from-csv-to-active-excel-sheet

    With wsTrack.QueryTables.Add(Connection:="TEXT;" & sTrackName, Destination:=wsTrack.Range("A1"))
         .TextFileParseType = xlDelimited
         .TextFileCommaDelimiter = True
         .Refresh
    End With
    
    TidyProcessedNav
    iTrackLastRow = FLastRow ' Small hack based on the fact I know ForceFindExtents is called inside TidyProcessedNav
    
    ' Now to populate the QC worksheet
    wsEvents.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    wsQC.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    wsWorking.Select
    Rows("2:" & iWorkingLastRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    wsQC.Select
    ForceFindExtents
    iKPCol = FindColumn("KP")
    
    Cells(FLastRow + 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Color the newly added rows yellow (to help identify the rows that will need inserting into VW)
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    FLastRow = FLastRow + iWorkingLastRow - 1
    
    TidyVWExcelExport
    
    ' Sort the resulting data by KP
    wsQC.AutoFilter.Sort.SortFields.Clear
    wsQC.AutoFilter.Sort.SortFields.Add Key:=Range(Cells(2, iKPCol), Cells(FLastRow, iKPCol)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    
    With wsQC.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Add a QC Column (distance from last event)
    Columns(iKPCol + 1).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Cells(1, iKPCol + 1).Value = "QC"
    Cells(3, iKPCol + 1).FormulaR1C1 = "=1000*(RC[-1]-R[-1]C[-1])"
    Cells(3, iKPCol + 1).Select
    Selection.Copy
    Range(Cells(3, iKPCol + 1), Cells(FLastRow, iKPCol + 1)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Columns(iKPCol + 1).Select
    Selection.NumberFormat = "0"
    
    wbCurrent.Save
    
    ' At this stage, we have our list of events to be imported, we have a QC spreadsheet to confirm if we need to make any changes to the underlying datasets
    ' and we also have our processed data ready for lookup
    
    wsWorking.Select
    
    ' Need to change this depending on inspection direction
    LookupDateTimeFromTrack (FMatchCode)
    
    wbCurrent.Save
    wsWorking.Activate
    wsWorking.SaveAs filename:=Text_Replace(sHistoricalName, ".xlsx", " - For Import.csv"), FileFormat:=xlCSV, Local:=True
    
    Application.DisplayAlerts = False
    wbCurrent.SaveAs filename:=sNewName, FileFormat:=xlWorkbookDefault, Local:=True
    Application.DisplayAlerts = True
    
    MsgBox "Done.  Please check the QC tabsheet for possible errors before importing the csv file into VisualWorks"
End Sub

Sub LookupDateTimeFromTrack(AMatchTime As Long)
'
'   Requires a separate Sheet exists called "Track"
'   Requires that current Sheet has first two columns for Date and Time and 7th column is KP
'
'  AMAtchTime = 1 - Use data from First Pass
'  AMatchTime = -1 - Use data from Last Pass
'
    Dim wsCurrent As Worksheet
    Dim iDateCol As Long, iTimeCol As Long
    
    Set wsCurrent = ActiveWorkbook.ActiveSheet
    
    wsCurrent.Activate
    
    ForceFindExtents
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    
    Columns(iDateCol).Select
    Selection.NumberFormat = "dd/mm/yyyy"
    Columns(iTimeCol).Select
    Selection.NumberFormat = "HH:mm:ss"
    
    Cells(2, 1).FormulaR1C1 = "=TEXT((RC[2]+RC[3])-TIME(8, 0, 0),""YYYYMMDDHHmmss"")&""000"""
    Cells(2, 2).FormulaR1C1 = "F"
    Cells(2, iDateCol).FormulaR1C1 = "=INDEX(Track!C[-2], MATCH(RC[6], Track!C[2], " & AMatchTime & "))"
    Cells(2, iTimeCol).FormulaR1C1 = "=INDEX(Track!C[-2], MATCH(RC[5], Track!C[1], " & AMatchTime & "))"
    
    Range(Cells(2, 1), Cells(2, iTimeCol)).Select
    Selection.Copy
    
    Range(Cells(2, 1), Cells(FLastRow, iTimeCol)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("A1").Select
End Sub

Function IncidentType(AIncidentTypeCode As String) As String
    If AIncidentTypeCode = "AN" Then
        IncidentType = "Anode"
    ElseIf AIncidentTypeCode = "FJ" Then
        IncidentType = "Field Joint"
    Else
        IncidentType = "ERROR"
    End If
End Function

Function Incident(AIncidentTypeCode As String, AIncidentCode As String) As String
    If AIncidentTypeCode = "AN" Then
        If AIncidentCode = "WTN" Then
            Incident = "Negligible Wastage <25%"
        Else
            Incident = "TODO"
        End If
    ElseIf AIncidentTypeCode = "FJ" Then
        If AIncidentCode = "GDC" Then
            Incident = "Good Condition"
        Else
            Incident = "TODO"
        End If
    Else
        Incident = "ERROR"
    End If
End Function

Sub TidyProcessedNav()
    Dim iDateCol As Long
    Dim iTimeCol As Long
    
    Call BasicTidy(ActiveSheet)
    
    Columns("A:A").Select
    Selection.NumberFormat = "dd/mm/yyyy"
    Columns("B:B").Select
    Selection.NumberFormat = "HH:mm:ss.000"
    Columns("C:D").Select
    Selection.NumberFormat = "0.00"
    Columns("E:E").Select
    Selection.NumberFormat = "0.0000"
    Columns("F:I").Select
    Selection.NumberFormat = "0.00"
    Range("D10").Select
    
    ForceFindExtents
    
    iDateCol = FindColumn("Date")
    iTimeCol = FindColumn("Time")
    
    ' Set to True if you want the Nav to be sorted
    If False Then
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
    End If
    
    Range("A2").Select
End Sub


