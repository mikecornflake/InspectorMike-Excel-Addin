Attribute VB_Name = "mod_NIS_Nexus5"
Option Explicit
Option Private Module

Dim FMultimedia As Worksheet
Dim FmmFilenameCol As Long, FmmNewFilenameCol As Long, FmmFolderCol As Long


Sub Nexus5_Build_Full_Name()
' Attribute Build_Full_Name.VB_Description = "Macro recorded 31/05/2007 by hutchida"
    Dim iLastRow As Integer
    Dim iLastColumn As Integer
    Dim i, iCol As Integer
    Dim sFormula As String
    
    ' Find the last Component.Location Column
    iCol = Find_Column("Workpack.Name") - 2
    
    ' Delete the Nexen / Buzzard Columns
    ' Columns("A:B").Select
    ' Range("B1").Activate
    ' Selection.Delete Shift:=xlToLeft
    ' Range("A2").Select
    
    ' Find last row
    Range("A1").Select
    Selection.End(xlDown).Select
    iLastRow = ActiveCell.Row
    
    ' Find Last Column
    Range("A1").Select
    Selection.End(xlToRight).Select
    iLastColumn = ActiveCell.Column
    
    Application.CutCopyMode = False
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
    ' Build up the formula
    sFormula = "="
    For i = 1 To iCol - 1
      sFormula = sFormula & "RC[" & i & "]&"" / ""&"
    Next i
    sFormula = sFormula & "RC[" & iCol & "]"
    
    Range("A2").Select
    ActiveCell.FormulaR1C1 = sFormula
    
    Range("A2").Select
    Selection.Copy
    
    Range("A" & iLastRow).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
    ' Populate the header
    Cells(1, 1).Value = "Location"
    
    ' Stop Col A from being a formula
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Columns("A:A").Select
    Selection.Replace What:="/  /  /  /  /  /  /  /  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/  /  /  /  /  /  /  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/  /  /  /  /  /  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/  /  /  /  /  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/  /  /  /  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/  /  /  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/  /  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="/  /", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    ' Shell Philippines
    Selection.Replace What:="Shell Philippines Exploration / Malampaya / ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range("A2").Select
    Application.CutCopyMode = False
End Sub

Public Sub Nexus5_Tidy_Event_Export()
    Dim iSheet As Long
    
    Call SortSheetsAlphabetically(ActiveWorkbook)

    If WorksheetExists("Findings") Then
        Sheets("Findings").Move Before:=Sheets(1)
    End If
    
    Find_Multimedia
    
    Delete_Lookup_Tabs
    
    For iSheet = 1 To Sheets.Count
        Sheets(iSheet).Select
        
        If Sheets(iSheet).Name <> "Multimedia" Then
            If Find_Column("Component.Location") > 0 Then
                ActiveSheet.Tab.ColorIndex = 43
                
                Remove_Formatting
                Remove_Conditional_Formatting
                Tidy_Header
                
                If Sheets(iSheet).Name <> "Findings" Then
                    Nexus5_Build_Full_Name
                Else
                    Rename_Column "Component.Location", "Location"
                End If
            
                Populate_Event_Name
                Tidy_Columns
                Process_Multimedia
                
                Cells.Select
                Cells.EntireColumn.AutoFit
                Cells.EntireRow.AutoFit
                
                Range("A2").Select
            Else
                ActiveSheet.Tab.Color = 15773696
            End If
        End If
    Next
    
    Hyperlink_Findings
    
    Leave_Sheets_Tidy
End Sub

Private Sub Process_Multimedia()
    Dim oCurrent As Worksheet
    Dim iMediaCol As Long, iColAdd As Long
    Dim sRootFolder As String, sImagesFolder As String
    Dim sNewLink As String
    Dim sOrigFilename As String
    Dim sNewFilename As String, sNewFolder As String
    Dim immRow As Long
    Dim iRow As Long
    
    ' Got an oddity with some files being reported as "Not Found" when they're clearly there.
    ' This will allow those files to be manually processed
    On Error Resume Next
    
    If FMultimedia Is Nothing Then
        Exit Sub
    End If
    
    Set oCurrent = ActiveWorkbook.ActiveSheet
    iMediaCol = Find_Column("Event.Multimedia")
    
    sImagesFolder = Path_GetFileNameNoExt(ActiveWorkbookLocalFilename) & "_Images\"
    sRootFolder = ActiveWorkbookPath & sImagesFolder
    
    If iMediaCol <= 0 Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    
    For iRow = 2 To oCurrent.UsedRange.Rows.Count
        iColAdd = 0
        
        While (Trim(oCurrent.Cells(iRow, iMediaCol + iColAdd).Value) <> "")
            sOrigFilename = Trim(oCurrent.Cells(iRow, iMediaCol + iColAdd).Value)
            
            immRow = FindInColumn(FMultimedia, FmmFilenameCol, sOrigFilename)
            
            If immRow > 0 Then
                sNewFolder = Trim(FMultimedia.Cells(immRow, FmmFolderCol).Value)
                
                sNewFilename = Format(immRow - 1, "0000") & " - " & File_SanitizeName(Trim(FMultimedia.Cells(immRow, FmmNewFilenameCol).Value))
                
                If sNewFolder = "DO NOT MOVE" Then
                    sNewLink = Replace(sImagesFolder & sNewFilename, " ", "%20")
                    If File_Exists(sRootFolder & sOrigFilename) Then
                        ' File_EnsureFolder (Path_AddTrailingDelimiter(sRootFolder & sNewFolder))
                        
                        Call FileCopy(sRootFolder & sOrigFilename, sRootFolder & sNewFilename)
                        If File_Exists(sRootFolder & sNewFilename) Then
                            Kill (sRootFolder & sOrigFilename)
                        End If
                    End If
                    
                    ' It's possible the image was already processed on a different tab, if so, this hyperlink still needs updating
                    If File_Exists(sRootFolder & sNewFilename) Then
                        oCurrent.Cells(iRow, iMediaCol + iColAdd).Select
                        Selection.Hyperlinks(1).Address = sNewLink
                        Selection.Hyperlinks(1).TextToDisplay = sNewFilename
                    End If
                Else
                    sNewLink = Replace(sImagesFolder & Path_AddTrailingDelimiter(sNewFolder) & sNewFilename, " ", "%20")
                    If File_Exists(sRootFolder & sOrigFilename) Then
                        File_EnsureFolder (Path_AddTrailingDelimiter(sRootFolder & sNewFolder))
                        
                        Call FileCopy(sRootFolder & sOrigFilename, sRootFolder & Path_AddTrailingDelimiter(sNewFolder) & sNewFilename)
                        If File_Exists(sRootFolder & Path_AddTrailingDelimiter(sNewFolder) & sNewFilename) Then
                            Kill (sRootFolder & sOrigFilename)
                        End If
                    End If
                    
                    ' It's possible the image was already processed on a different tab, if so, this hyperlink still needs updating
                    If File_Exists(sRootFolder & Path_AddTrailingDelimiter(sNewFolder) & sNewFilename) Then
                        oCurrent.Cells(iRow, iMediaCol + iColAdd).Select
                        Selection.Hyperlinks(1).Address = sNewLink
                        Selection.Hyperlinks(1).TextToDisplay = sNewFilename
                    End If
                End If
            End If
            
            iColAdd = iColAdd + 1
        Wend
    Next iRow
    
    If ActiveWorkbook.ActiveSheet.Name <> oCurrent.Name Then
        oCurrent.Select
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Sub Find_Multimedia()
    Dim iSheet As Long
    
    Set FMultimedia = Nothing
    
    For iSheet = 1 To ActiveWorkbook.Sheets.Count
        If Sheets(iSheet).Name = "Multimedia" Then
            Set FMultimedia = Sheets(iSheet)
            
            FMultimedia.Select
            
            FmmFilenameCol = Find_Column("Filename")
            FmmNewFilenameCol = Find_Column("New_Filename")
            FmmFolderCol = Find_Column("Recording_Folder")
        End If
    Next iSheet
End Sub

Private Sub Delete_Lookup_Tabs()
    Dim iSheet As Long
    
    Application.DisplayAlerts = False
    For iSheet = Sheets.Count To 1 Step -1
        If Trim(Sheets(iSheet).Cells(1, 2).Value) = "" Then
            Sheets(iSheet).Delete
        End If
    Next iSheet
    Application.DisplayAlerts = True
End Sub


Private Sub Leave_Sheets_Tidy()
    Dim i As Long
    Dim iFinding, iMultimedia As Long
    
    iMultimedia = -1
    iFinding = -1
    
    For i = 1 To Sheets.Count
        Sheets(i).Select
        
        If Sheets(i).Name = "Multimedia" Then
            iMultimedia = i
        ElseIf Sheets(i).Name = "Finding" Then
            iFinding = i
        End If
    
        Cells.Select
        Selection.Font.Name = "Tahoma"
        Selection.Font.size = 10
        
        Cells(2, 1).Select
        Cells(1, 1).Select
    Next i
    
    If iFinding <> -1 Then
        Sheets(iFinding).Select
    ElseIf Sheets(1).Name <> "Multimedia" Then
        Sheets(1).Select
    Else
        Sheets(2).Select
    End If
    
    If iMultimedia <> -1 Then
        Application.DisplayAlerts = False
        Sheets(iMultimedia).Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub Hyperlink_Findings()
    '
    ' The idea behind this is every row in Findings will contain a hyperlink to the
    ' approriate Row in the relevant event sheet
    '
    
    Dim iLastFinding, iFindingEventCol As Long
    Dim iEventCol As Long
    
    Dim iFinding, iEvent As Long
    Dim sEvent As String
    Dim sFullEvent As String
        
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    
    If WorksheetExists("Findings") Then
        Sheets("Findings").Select
        ActiveSheet.Tab.ColorIndex = 22
        
        If Trim(Cells(2, 1).Value) <> "" Then
            ForceFindExtents
            
            iLastFinding = FLastRow
            iFindingEventCol = FindColumn("Event")
            
            For iFinding = 2 To iLastFinding
                Sheets("Findings").Select
                
                sFullEvent = Cells(iFinding, iFindingEventCol)
                sEvent = Left(sFullEvent, Text_FindLast(sFullEvent, " ") - 1)
                
                If WorksheetExists(sEvent) Then
                    Sheets(sEvent).Select
                    ActiveSheet.Tab.ColorIndex = 22
                    
                    ForceFindExtents
                    
                    iEventCol = FindColumn("Event")
                    
                    If iEventCol <> 0 Then
                        iEvent = FindInColumn(ActiveSheet, iEventCol, sFullEvent)
                        
                        If iEvent <> 0 Then
                            ' Color the Event with the findings
                            Rows(iEvent).Select
                            With Selection.Interior
                                .ColorIndex = 22
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                            End With
                            
                            ' Put in a hyperlink to the Findings Page
                            Cells(iEvent, iEventCol).Select
                            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'Findings'!A" & iFinding, TextToDisplay:=sFullEvent
                            
                            ' On the Findings Sheet, put in the hyperlink to the Events Page
                            Sheets("Findings").Select
                            Cells(iFinding, iFindingEventCol).Select
                            
                            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & sEvent & "'!A" & iEvent, TextToDisplay:=sFullEvent
                        End If
                    End If
                End If
            Next iFinding
        End If
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Sub Tidy_Header()
    Dim iStart, iEnd As Long
    Dim i As Long
    
    ' Reduce to one column
    Range("A2").Select
    i = 1
    
    While Trim(Cells(2, i).Value) = ""
        ' Cells(2, i).Select
        i = i + 1
    Wend
    
    While Trim(Cells(2, i).Value) <> ""
        ' Cells(1, i).Select
        Cells(1, i).Value = Cells(2, i).Value & "." & Cells(3, i).Value
        i = i + 1
    Wend
    
    ForceFindExtents
    Range(Cells(2, 1), Cells(3, FLastColumn)).Select
    Selection.Delete Shift:=xlUp
    
    ' Prettify the head row
    Rows("1:1").Select
    Selection.Interior.ColorIndex = 40
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ' Get rid of the obscure column names
    Rows("1:1").Select
    Range("B1").Activate
    Selection.Replace What:=".Event", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("A2").Select
End Sub

Private Sub Populate_Event_Name()
    Dim iEventName, iEventNumber As Long
    Dim iLastRow As Long
    Dim i As Long
    
    iEventName = Find_Column("Event.Name")
    iEventNumber = Find_Column("Event.Number")
    
    ' Find Last Row
    Range("A1").Select
    
    If Trim(Cells(2, 1).Value) <> "" Then
        Selection.End(xlDown).Select
        iLastRow = ActiveCell.Row
    
        If (iEventName <> -1) And (iEventNumber <> -1) Then
            For i = 2 To iLastRow
                Cells(i, iEventName).Value = Cells(i, iEventName).Value & " " & Cells(i, iEventNumber).Value
            Next
        End If
    End If
End Sub

Private Sub Tidy_Columns()
    Dim iLastRow, iLastColumn As Long
    Dim iRow, iCol As Long
    Dim bEmpty As Boolean
    
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    
    While Delete_Column("Component.Location")
    Wend

    Delete_Column ("Workpack.Name")

    Delete_Column ("Event.Number")
    Delete_Column ("Video Counter.Survey Start")
    Delete_Column ("Video Counter.Survey End")
    
    'Delete_Column ("Easting.Survey Start")
    'Delete_Column ("Easting.Survey End")
    
    'Delete_Column ("Northing.Survey Start")
    'Delete_Column ("Northing.Survey End")
    
    Delete_Column ("Photo?.Survey Start")
    Delete_Column ("Photo?.Survey End")
    
    Delete_Column ("Spare3.Survey Start")
    Delete_Column ("Spare4.Survey Start")
    Delete_Column ("Spare3.Survey End")
    Delete_Column ("Spare4.Survey End")
    
    Rename_Column "Event.Name", "Event"
    Rename_Column "Event.Comments", "Comments"
    Rename_Column "Clock.Survey Start", "Start Time"
    Rename_Column "Clock.Survey End", "End Time"
    Rename_Column "KP.Survey Start", "Start KP"
    Rename_Column "KP.Survey End", "End KP"
    Rename_Column "Easting.Survey Start", "Start Easting"
    Rename_Column "Easting.Survey End", "End Easting"
    Rename_Column "Northing.Survey Start", "Start Northing"
    Rename_Column "Northing.Survey End", "End Northing"
    Rename_Column "Offset.Survey Start", "Start Offset"
    Rename_Column "Offset.Survey End", "End Offset"
    Rename_Column "Depth.Survey Start", "Start Depth"
    Rename_Column "Depth.Survey End", "End Depth"
    Rename_Column "Temperature.Survey Start", "Start Temperature"
    Rename_Column "Temperature.Survey End", "End Temperature"
    Rename_Column "Heading.Survey Start", "Start Heading"
    Rename_Column "Heading.Survey End", "End Heading"
    
    Rename_Column ".", "Images"
    Rename_Column "Finding.Code", "Code"
    Rename_Column "Finding.ID", "ID"
    Rename_Column "Finding.Description", "Finding"
    
    ' Find last row
    If Trim(Cells(2, 1).Value) <> "" Then
        Range("A1").Select
        Selection.End(xlDown).Select
        iLastRow = ActiveCell.Row
        
        ' Find Last Column
        Range("A1").Select
        Selection.End(xlToRight).Select
        iLastColumn = ActiveCell.Column
        
        ' Delete the empty columns
        For iCol = iLastColumn To 3 Step -1
            bEmpty = True
            
            For iRow = 2 To iLastRow
                bEmpty = bEmpty And (Trim(Cells(iRow, iCol).Value) = "")
            Next iRow
            
            If bEmpty Then
                Range(Cells(1, iCol), Cells(iLastRow + 1, iCol)).Delete xlShiftToLeft
            End If
        Next iCol
        
        ' Delete End Time if it's not valid
        iCol = Find_Column("End Time")
        If iCol <> -1 Then
            bEmpty = True
            
            For iRow = 2 To iLastRow
                If (Trim(Cells(iRow, iCol).Value) = "00:00:00") Then
                    ' Clean up the cells in case there is a value in this row somewhere
                    Cells(iRow, iCol).Value = ""
                Else
                    bEmpty = False
                End If
            Next iRow
            
            If bEmpty Then
                Columns(iCol).Delete
            End If
        End If
    End If
                
    Range("A1").Select
    Selection.AutoFilter
    
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Application.ScreenUpdating = True
End Sub

Private Sub Remove_Formatting()
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders.LineStyle = xlNone
    Selection.Font.Bold = False
End Sub

Private Sub Remove_Conditional_Formatting()
    Cells.Select
    Cells.FormatConditions.Delete
End Sub
