Attribute VB_Name = "mod_Fugro_SenseStructures"
' ============================================================================
' Module: SenseStructuresExportProcessing
' Purpose: Process event data, normalize sheets,
'          link anomalies and multimedia
' Entry Points:
'     Public Sub ProcessSheet()        ## This is for most use cases
'     Public Sub HyperlinkImages(...)  ## Preparing for future use cases
'
' All other routines are Private
'
'   In Sense.Structures Review/Exporting, ensure you have a suitable XML loaded
'   i.e. "Inspector Mike - Task and Event.xml"
'
'   At the very least, you want Common, Event, Anomaly and multimedia fields.  I suggest the following:
'     Workpack    Component   Type    IncidentID  EventCode   SubEventCode    StartTime   EndTime
'     Start KP    End KP  Task/Event Length   Length (mm) Width (mm)  Height (mm) Easting Northing    Depth   Heading DCC
'     Location    Comment Anomaly Description
'     Image1  Image2  Image3  Video   Anomaly_Image1  Anomaly_Image2  Anomaly_Image3  Anomaly_Video
'     Active  Secure  Depletion
'     CP Value
'     Percentage Hard Percentage Soft Thickness Hard  Thickness Soft
'     Flooded
'
'   Using Sense.Structures Review/Exporting, ensure you have the correct workpacks selected
'   then export the data and save the Excel to your working folder
'
'   Copy all project multimedia to a subfolder (maybe not video)
'   and run Process()
'
'   You will be asked to select a media folder.
'
'   If the media folder is a subfolder of the Excel file - then hyperlinks will be relative, and the Excel file/Media can be distributed together
'   If the media folder isn't a subfolder, then the hyperlinks will be absolute, and the Excel file will only have limited distribution (same vessel or office)
'   If you click cancel on the Choose Folder dialog, then no media hyperlinks will be added
'
'   7 Aug 2025
'   Mike Thompson & copilot / Atlantis Dweller
'   No data was shared with copilot, only algorithm based requests. Their help/input was invaluble
'   mike.cornflake@gmail.com
'
'
' ============================================================================
'
'   Version history:
'
'     07/08/2025: Initial release developed for Ineos Structures campaign against Sense.Structures 1.3.5.9
'     08/08/2025: Add ForceRelative to image/video search.  Allows for media folders to be in different areas to Excel file.
'                 - Change Private Sub HyperlinkImages(baseFolder As String)
'                       To Private Sub HyperlinkImages(baseFolder As String, Optional forceRelative As Boolean = False)
'                 - Add Private Function GetRelativePath(fromPath As String, toPath As String) As String
'     09/08/2025: Allow HyperlinkImages() to be reCalled with a different media folder.  Clears all existing links, then adds new.
'
' ============================================================================

Public Sub ProcessSheet()
    If ActiveSheet.Cells(1, 1).Value <> "Workpack" Or ActiveSheet.Cells(1, 2).Value <> "Component" Then
        MsgBox ("Ensure you have the correct workbook loaded")
        Exit Sub
    End If
    
    ' Identify our original sheet
    ActiveSheet.Name = "Original"
    
    ' Delete unused columns by field name
    Application.StatusBar = "Deleting columns..."
    Call DeleteColumn("Type")
    Call DeleteColumn("IncidentID")
    Call DeleteColumn("SubEventCode")
    Call DeleteColumn("Start KP")
    Call DeleteColumn("End KP")
    Call DeleteColumn("Task/Event Length")
    Call DeleteColumn("Length (mm)")
    Call DeleteColumn("Width (mm)")
    Call DeleteColumn("Height (mm)")
    Call DeleteColumn("DCC")

    ' Split data into event-specific sheets
    Application.StatusBar = "Normalising events into worksheets..."
    Call SplitByEventCode

    ' Sort sheets alphabetically
    Call SortSheetsAlphabetically
    
    ' Copy all the anomalies into a new sheet
    Application.StatusBar = "Processing anomalies..."
    BuildAnomalySheet
    
    ' Delete existing sheet
    DeleteSheet ("Original")
    
    ' Hyperlink the media
    Dim mediaFolder As String
    
    Application.StatusBar = "Hyperlinking multimedia..."
    mediaFolder = PickFolder(ActiveWorkbook.Path, "Select the project folder containing multimedia...")
    If mediaFolder <> "" Then
        Call HyperlinkImages(mediaFolder, True)
    Else
        MsgBox "No folder selected. Multimedia linking skipped.", vbInformation
    End If
    
    ' Deleting unused columns
    Application.StatusBar = "Deleting unused columns..."
    DeleteEmptyColumns

    ' Format all sheets
    Call FormatAllSheets
    
    ' Select the first sheet
    ActiveWorkbook.Sheets(1).Activate
    
    Application.StatusBar = False
End Sub

Private Sub FormatTable()
'
' FormatTable Macro
'
'
    Dim colIndex As Long
    
    Range("A1").Select
    
    ' Header
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    ' Filter
    Range("A1").Select
    If Not ActiveSheet.AutoFilterMode Then
        Selection.AutoFilter
    End If
    
    ' Get the font right
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .size = 9
    End With
    With Selection
        .VerticalAlignment = xlTop
    End With
    
    ' Autosize
    Cells.Select
    Selection.ColumnWidth = 50
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
    Set ws = ActiveSheet
    colIndex = 1

    ' Loop until the first row cell is empty
    Do While Trim(ws.Cells(1, colIndex).Value) <> ""
        If ws.Columns(colIndex).ColumnWidth > 65 Then
            ws.Columns(colIndex).ColumnWidth = 65
        End If
        colIndex = colIndex + 1
    Loop
    
    ' Home
    Range("A1").Select
End Sub

Private Sub RemoveExternalHyperlinks()
    Dim ws As Worksheet
    Dim hl As Hyperlink
    Dim i As Long
    Dim rng As Range
    Dim cellFormat As Variant

    For Each ws In ActiveWorkbook.Worksheets
        For i = ws.Hyperlinks.Count To 1 Step -1
            Set hl = ws.Hyperlinks(i)
            Set rng = hl.Range
            
            ' Check if hyperlink is external
            If Not hl.Address Like "#*" Then
                If hl.Address Like "*.jpg" Or _
                   hl.Address Like "*.jpeg" Or _
                   hl.Address Like "*.png" Or _
                   hl.Address Like "*.mp4" Or _
                   hl.Address Like "http*" Or _
                   hl.Address Like "https*" Or _
                   hl.Address Like "www.*" Then
                   
                    ' Save cell formatting
                    cellFormat = GetCellFormat(rng)
                    
                    ' Remove hyperlink
                    hl.Delete
                    
                    ' Restore formatting
                    ApplyCellFormat rng, cellFormat
                    
                    ' And undo the solid white that hyperlinks have
                    If rng.Interior.ColorIndex = 2 Then
                        rng.Interior.ColorIndex = xlColorIndexNone
                    End If
                End If
            End If
        Next i
    Next ws
End Sub

' Helper: Capture cell formatting
Private Function GetCellFormat(rng As Range) As Variant
    Dim fmt(1 To 5) As Variant
    fmt(1) = rng.Interior.Color
    fmt(2) = rng.Font.Name
    fmt(3) = rng.Font.size
    fmt(4) = rng.HorizontalAlignment
    fmt(5) = rng.VerticalAlignment
    GetCellFormat = fmt
End Function

' Helper: Apply formatting back to cell
Private Sub ApplyCellFormat(rng As Range, fmt As Variant)
    rng.Interior.Color = fmt(1)
    rng.Font.Name = fmt(2)
    rng.Font.size = fmt(3)
    rng.HorizontalAlignment = fmt(4)
    rng.VerticalAlignment = fmt(5)
End Sub


Private Function DeleteColumn(fieldName As String) As Boolean
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim cell As Range
    Dim colIndex As Long
    
    ' Use the active sheet, or change to a specific sheet if needed
    Set ws = ActiveSheet
    Set headerRow = ws.Rows(1)
    
    ' Loop through each cell in the header row
    For Each cell In headerRow.Cells
        If LCase(Trim(cell.Value)) = LCase(fieldName) Then
            colIndex = cell.Column
            ws.Columns(colIndex).Delete
            
            DeleteColumn = True
            
            Exit Function
        End If
    Next cell
    
    DeleteColumn = False
End Function

Private Function JoinArrays(arr1 As Variant, arr2 As Variant) As Variant
    Dim result() As Variant
    Dim i As Long, total As Long

    total = UBound(arr1) + UBound(arr2) + 2
    ReDim result(0 To total - 1)

    For i = 0 To UBound(arr1)
        result(i) = arr1(i)
    Next i

    For i = 0 To UBound(arr2)
        result(UBound(arr1) + 1 + i) = arr2(i)
    Next i

    JoinArrays = result
End Function

Private Sub SplitByEventCode()
    Dim wsSource As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim headerRow As Range
    Dim colMap As Object
    Dim eventSheets As Object
    Dim eventFields As Object
    Dim coreFields As Variant
    Dim r As Long
    Dim eventCode As String
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim fieldList As Variant
    Dim f As Variant

    Set wsSource = ActiveSheet
    Set headerRow = wsSource.Rows(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    ' Map column headers to their column numbers
    Set colMap = CreateObject("Scripting.Dictionary")
    For Each cell In headerRow.Cells
        If Trim(cell.Value) <> "" Then
            colMap(LCase(Trim(cell.Value))) = cell.Column
        End If
    Next cell

    ' Define core fields
    coreFields = Array("Workpack", "Component", "EventCode", "StartTime", "EndTime", _
                       "Easting", "Northing", "Depth", "Heading", "Location", "Comment", _
                       "Image1", "Image2", "Image3", "Anomaly", "Description", "Video", _
                       "Anomaly_Image1", "Anomaly_Image2", "Anomaly_Image3", "Anomaly_Video")

    ' Define event-specific fields
    Set eventFields = CreateObject("Scripting.Dictionary")
    eventFields.Add "AW", Array("Active", "Secure", "Depletion")
    eventFields.Add "CP-PROX", Array("CP Value")
    eventFields.Add "CP-CON", Array("CP Value")
    eventFields.Add "MG", Array("Percentage Hard", "Percentage Soft", "Thickness Hard", "Thickness Soft")
    eventFields.Add "FMD", Array("Flooded")

    ' Track created sheets
    Set eventSheets = CreateObject("Scripting.Dictionary")

    ' Loop through each row
    For r = 2 To lastRow
        eventCode = Trim(wsSource.Cells(r, colMap("eventcode")).Value)
        If eventCode = "" Then GoTo NextRow

        ' Create sheet if it doesn't exist
        If Not eventSheets.Exists(eventCode) Then
            Set targetSheet = Sheets.Add(After:=Sheets(Sheets.Count))
            targetSheet.Name = eventCode
            eventSheets.Add eventCode, targetSheet

            ' Write headers
            fieldList = coreFields
            If eventFields.Exists(eventCode) Then
                fieldList = JoinArrays(coreFields, eventFields(eventCode))
            End If

            For i = 0 To UBound(fieldList)
                targetSheet.Cells(1, i + 1).Value = fieldList(i)
            Next i
        Else
            Set targetSheet = eventSheets(eventCode)
        End If

        ' Write data
        targetRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
        fieldList = coreFields
        If eventFields.Exists(eventCode) Then
            fieldList = JoinArrays(coreFields, eventFields(eventCode))
        End If

        For i = 0 To UBound(fieldList)
            f = LCase(fieldList(i))
            If colMap.Exists(f) Then
                targetSheet.Cells(targetRow, i + 1).Value = wsSource.Cells(r, colMap(f)).Value
            End If
        Next i

NextRow:
    Next r
End Sub

Private Sub SortSheetsAlphabetically()
    Dim i As Long, j As Long
    Dim tempName As String

    ' Bubble sort
    For i = 1 To Sheets.Count - 1
        For j = i + 1 To Sheets.Count
            If Sheets(i).Name > Sheets(j).Name Then
                Sheets(j).Move Before:=Sheets(i)
            End If
        Next j
    Next i
End Sub

Private Sub FormatAllSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ws.Select
        
        Call FormatTable
        
        Call ColumnLimitWidth("Comment", 30)
        Call ColumnLimitWidth("Component", 30)
    Next ws
End Sub

Private Sub DeleteSheet(sheetName As String)
    Set wb = ActiveWorkbook
    ' Delete existing sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False
End Function

Private Sub BuildAnomalySheet()
    Dim wb As Workbook
    Dim wsAnomaly As Worksheet
    Dim ws As Worksheet
    Dim colMap As Object
    Dim anomalyHeaders As Variant
    Dim lastRow As Long, targetRow As Long
    Dim r As Long, i As Long
    Dim cell As Range
    Dim f As Variant
    Dim sourceRowAddress As String
    Dim hasAnomaly As Boolean

    Set wb = ActiveWorkbook

    ' Delete existing sheet if it exists
    DeleteSheet ("Anomaly")

    ' Create new Anomaly sheet
    Set wsAnomaly = wb.Sheets.Add(Before:=wb.Sheets(1))
    wsAnomaly.Name = "Anomaly"
    wsAnomaly.Tab.Color = RGB(255, 0, 0)
                
    ' Define headers
    anomalyHeaders = Array("Workpack", "Component", "EventCode", "StartTime", "Easting", "Northing", _
                           "Depth", "Comment", "Anomaly", "Description", "Anomaly_Image1", _
                           "Anomaly_Image2", "Anomaly_Image3", "Anomaly_Video")

    ' Write headers
    For i = 0 To UBound(anomalyHeaders)
        wsAnomaly.Cells(1, i + 1).Value = anomalyHeaders(i)
    Next i

    targetRow = 2 ' Start writing from row 2

    ' Loop through all sheets except Anomaly
    For Each ws In wb.Sheets
        If (ws.Name <> "Anomaly" And ws.Name <> "Original") Then
            hasAnomaly = False
        
            ' Map column headers
            Set colMap = CreateObject("Scripting.Dictionary")
            For Each cell In ws.Rows(1).Cells
                If Trim(cell.Value) <> "" Then
                    colMap(LCase(Trim(cell.Value))) = cell.Column
                End If
            Next cell

            ' Skip if "Anomaly" column doesn't exist
            If Not colMap.Exists("anomaly") Then GoTo NextSheet

            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            ' Loop through rows
            For r = 2 To lastRow
                If Trim(ws.Cells(r, colMap("anomaly")).Value) <> "" Then
                    hasAnomaly = True
                    
                    ' Copy relevant columns to Anomaly sheet
                    For i = 0 To UBound(anomalyHeaders)
                        f = LCase(anomalyHeaders(i))
                        If colMap.Exists(f) Then
                            wsAnomaly.Cells(targetRow, i + 1).Value = ws.Cells(r, colMap(f)).Value
                        End If
                    Next i

                    ' Add hyperlink in EventCode cell to Anomaly sheet
                    sourceRowAddress = wsAnomaly.Cells(targetRow, 3).Address(External:=True)
                    ws.Hyperlinks.Add Anchor:=ws.Cells(r, colMap("eventcode")), _
                        Address:="", SubAddress:="'" & wsAnomaly.Name & "'!" & wsAnomaly.Cells(targetRow, 1).Address, _
                        TextToDisplay:=ws.Cells(r, colMap("eventcode")).Value

                    ' Add hyperlink in Anomaly sheet back to source row
                    wsAnomaly.Hyperlinks.Add Anchor:=wsAnomaly.Cells(targetRow, 3), _
                        Address:="", SubAddress:="'" & ws.Name & "'!" & ws.Cells(r, 1).Address, _
                        TextToDisplay:=ws.Cells(r, colMap("eventcode")).Value

                    ' Highlight the entire row in the event sheet
                    ws.Rows(r).Interior.Color = RGB(255, 200, 200)
                    
                    targetRow = targetRow + 1
                End If
            Next r
            
            ' If anomalies were found, color the sheet tab red
            If hasAnomaly Then
                ws.Tab.Color = RGB(255, 0, 0)
            End If
        End If
NextSheet:
    Next ws

    ' Format Anomaly sheet
    wsAnomaly.Rows(1).AutoFilter
    wsAnomaly.Rows(1).Interior.Color = RGB(255, 230, 230)
    wsAnomaly.Activate
    ActiveWindow.FreezePanes = False
    wsAnomaly.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Function GetRelativePath(fromPath As String, toPath As String) As String
    Dim fromParts() As String, toParts() As String
    Dim i As Long, commonIndex As Long
    Dim relPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    fromParts = Split(fso.GetParentFolderName(fromPath), "\")
    toParts = Split(fso.GetParentFolderName(toPath), "\")

    ' Find common root
    For i = 0 To Application.min(UBound(fromParts), UBound(toParts))
        If LCase(fromParts(i)) <> LCase(toParts(i)) Then Exit For
        commonIndex = i
    Next i

    ' Climb up from source
    For i = commonIndex + 1 To UBound(fromParts)
        relPath = relPath & "..\"
    Next i

    ' Descend to target
    For i = commonIndex + 1 To UBound(toParts)
        relPath = relPath & toParts(i) & "\"
    Next i

    ' Add filename
    relPath = relPath & fso.GetFileName(toPath)

    GetRelativePath = relPath
End Function

Private Sub Test()
    Call HyperlinkImages("Z:\505922_INEOS_FPS\Deliverable\GVI_DVI_WT_ACFM", True)
    ' Call HyperlinkImages("Z:\505922_INEOS_FPS\Deliverable\FMD", True)
    
End Sub

Public Sub HyperlinkImages(baseFolder As String, Optional forceRelative As Boolean = False)
    Dim fso As Object, fileDict As Object
    Dim wbPath As String, useRelative As Boolean
    Dim file As Object, folder As Object
    Dim ws As Worksheet, rng As Range
    Dim cell As Range, rowRange As Range
    Dim targetCols As Variant, colIndex As Long
    Dim filePath As String, relPath As String
    Dim filename As String
    Dim wbDrive As String, mediaDrive As String
    Dim FontSize As Single

    ' Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileDict = CreateObject("Scripting.Dictionary")
    
    wbPath = ActiveWorkbook.FullName
    wbDrive = Left(wbPath, 2)
    mediaDrive = Left(baseFolder, 2)

    ' Determine hyperlinking mode
    If forceRelative Then
        useRelative = (wbDrive = mediaDrive)
    Else
        useRelative = (InStr(1, baseFolder, ActiveWorkbook.Path, vbTextCompare) = 1)
    End If

    ' Recursively collect files
    Call CollectMediaFiles(fso.GetFolder(baseFolder), fileDict)

    ' Target columns to hyperlink
    targetCols = Array("Image1", "Image2", "Image3", "Video", _
                       "Anomaly_Image1", "Anomaly_Image2", "Anomaly_Image3", "Anomaly_Video")

    ' Loop through worksheets
    For Each ws In ActiveWorkbook.Worksheets
        Application.StatusBar = "Hyperlinking multimedia in " & ws.Name & "..."
        
        ' First, delete any existing external links
        Call RemoveExternalHyperlinks
        
        ws.Activate
    
        ' Find header row (assumes headers in row 1)
        Set rng = ws.Rows(1)
        For Each cell In rng.Cells
            For colIndex = LBound(targetCols) To UBound(targetCols)
                If Trim(cell.Value) = targetCols(colIndex) Then
                    ' Loop through data rows
                    For Each rowRange In ws.Range(cell.Offset(1, 0), ws.Cells(ws.Rows.Count, cell.Column).End(xlUp)).Rows
                        Dim dataCell As Range
                        Set dataCell = ws.Cells(rowRange.Row, cell.Column)
    
                        If Len(dataCell.Value) > 0 Then
                            filename = Trim(dataCell.Value)
                            If fileDict.Exists(filename) Then
                                filePath = fileDict(filename)
                                FontSize = dataCell.Font.size
                                If useRelative Then
                                    relPath = GetRelativePath(wbPath, filePath)
                                    dataCell.Hyperlinks.Add Anchor:=dataCell, Address:=relPath, TextToDisplay:=filename
                                Else
                                    dataCell.Hyperlinks.Add Anchor:=dataCell, Address:=filePath, TextToDisplay:=filename
                                End If
                                ' Setting a hyperlink appears to reset the font size.  This should preserve it
                                dataCell.Font.size = FontSize
                            End If
                        End If
                    Next rowRange
                End If
            Next colIndex
        Next cell
    Next ws
    
    Application.StatusBar = False
End Sub

Private Sub CollectMediaFiles(folder As Object, fileDict As Object)
    Dim fso As Object
    Dim file As Object, subFolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "jpg" Or LCase(fso.GetExtensionName(file.Name)) = "mp4" Or LCase(fso.GetExtensionName(file.Name)) = "png" Then
            If Not fileDict.Exists(file.Name) Then
                fileDict.Add file.Name, file.Path
            End If
        End If
    Next file
    
    For Each subFolder In folder.SubFolders
        Application.StatusBar = "Searching for media " & subFolder.Path & "..."
        CollectMediaFiles subFolder, fileDict
    Next subFolder
End Sub

Private Sub DeleteEmptyColumns()
    Dim ws As Worksheet
    Dim col As Long
    Dim lastRow As Long
    Dim isEmpty As Boolean

    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        
        With ws
            ' Start from last column and move left to avoid shifting issues
            For col = .Cells(1, .Columns.Count).End(xlToLeft).Column To 1 Step -1
                lastRow = .Cells(.Rows.Count, col).End(xlUp).Row
                If lastRow = 1 And Len(.Cells(1, col).Value) > 0 Then
                    .Columns(col).Delete
                End If
            Next col
        End With
    Next ws
End Sub

Private Sub ColumnLimitWidth(colName As String, colWidth As Double)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim colNum As Long

    For Each ws In ActiveWorkbook.Worksheets
        Set rng = ws.Rows(1)
        For Each cell In rng.Cells
            If Trim(cell.Value) = colName Then
                colNum = cell.Column
                With ws.Columns(colNum)
                    .WrapText = True
                    .ColumnWidth = colWidth
                End With
                Exit For
            End If
        Next cell
    Next ws
End Sub

Private Function PickFolder(baseFolder As String, Optional prompt As String = "Select a folder") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = prompt
        .InitialFileName = baseFolder
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = ""
        End If
    End With
End Function
