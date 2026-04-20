Attribute VB_Name = "mod_IntegrityElements"
' 2025 08 16 - Merged code written for Fugro/Apache code written for DOF/Chevron

Option Explicit

'
'  Input is the Workpack worksheet exported by IntegrityElements
'
Public Sub FormatIEWorkpackDataSheet()
'
    Dim oSheet As Worksheet
    Dim iCol As Long, iRow As Long, iMaxRow As Long
    Dim sColumn As String, sValue As String
    Dim Value
    
    For Each oSheet In ActiveWorkbook.Sheets
        oSheet.Activate
        
        Range("B1").Select
        
        ' Is this a likely workpack sheet?
        If ActiveCell.Value = "Substructure" Then
            ' Apply the filter
            Selection.AutoFilter
            
            iCol = 2
            ' For each Column
            While Trim(Cells(1, iCol).Value) <> ""
                ' Resize the columns
                Columns(iCol).Select
                Columns(iCol).EntireColumn.AutoFit
                
                ' Shade the header
                Cells(1, iCol).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.149998474074526
                    .PatternTintAndShade = 0
                End With
                
                ' Find the Row Count
                iRow = 2
                While Trim(Cells(iRow, 2).Value) <> ""
                    iRow = iRow + 1
                Wend
                iMaxRow = iRow - 1
                
                ' Do we need to convert numerical data to numbers?
                sColumn = LCase(Trim(Cells(1, iCol).Value))
                If (sColumn = "reading") Or (sColumn = "min") Or (sColumn = "max") Or (sColumn = "% hard") Or (sColumn = "mm hard") Or (sColumn = "% soft") Or (sColumn = "mm soft") Or (sColumn = "heading") Or (sColumn = "easting") Or (sColumn = "northing") Or (sColumn = "depth (m) rov") Then
                    iRow = 2
                    
                    While iRow <= iMaxRow
                        sValue = Trim(Cells(iRow, iCol))
                        Value = Val(sValue)
                        
                        If (sValue <> "") Then Cells(iRow, iCol).Value = Value
                        iRow = iRow + 1
                    Wend
                End If
                
                ' Goto next column
                iCol = iCol + 1
            Wend
            
            ' Sort the Sheet by Substructure
            ActiveSheet.AutoFilter.Sort.SortFields.Clear
            ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range("B1:B" & iMaxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range("C1:C" & iMaxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range("D1:D" & iMaxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveSheet.AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            ' Freeze the first row
            With ActiveWindow
                .SplitColumn = 0
                .SplitRow = 1
            End With
            ActiveWindow.FreezePanes = True
        
            Range("B1").Select
            
            ' Rename the worksheet
            sCaption = oSheet.Name
            sCaption = Trim(Mid(sCaption, 1, InStr(sCaption, " ")))
            If sCaption <> "" Then
                oSheet.Name = sCaption
            End If
        End If
    Next oSheet
    
    SortSheetsAlphabetically
    
    ActiveWorkbook.Sheets(1).Activate
End Sub

Private Sub SortSheetsAlphabetically()
    Dim i As Integer, j As Integer
    Dim tempSheet As Object
    
    ' Loop through all sheets in the workbook
    For i = 1 To ActiveWorkbook.Sheets.Count - 1
        For j = i + 1 To ActiveWorkbook.Sheets.Count
            ' Compare the names of adjacent sheets
            If ActiveWorkbook.Sheets(j).Name < ActiveWorkbook.Sheets(i).Name Then
                ' Swap sheets by moving the latter before the former
                Set tempSheet = ActiveWorkbook.Sheets(j)
                tempSheet.Move Before:=ActiveWorkbook.Sheets(i)
            End If
        Next j
    Next i
End Sub

Private Sub Format_Table()
'
' Format_Table Macro
'

'
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("D2").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Rows("1:1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1").Select
End Sub

Private Function NormalizeText(txt As String) As String
    Dim clean As String
    clean = Trim(txt)
    clean = Replace(clean, Chr(160), " ") ' non-breaking space
    ' clean = Replace(clean, Chr(8203), "")  ' zero-width space
    clean = Replace(clean, vbCrLf, " ")
    clean = Replace(clean, vbLf, " ")
    NormalizeText = clean
End Function


Private Sub PopulateAnomalyReferences()
    Dim wsCompletion As Worksheet
    Dim lastRowCompletion As Long
    Dim dict As Object
    Dim i As Long, lastCol As Long
    Dim asset As String
    Dim anomalyRef As String
    Dim sourceSheets As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Sheets containing anomalies
    sourceSheets = Array("Pipeline Anomalies", "Structure Anomalies")
    
    ' Set destination sheet
    Set wsCompletion = ActiveWorkbook.Sheets("Completion Report")
    If wsCompletion.AutoFilterMode Then wsCompletion.AutoFilterMode = False
    
    Set dict = CreateObject("Scripting.Dictionary")

    ' Loop through each source sheet
    For Each sheetName In sourceSheets
        Set ws = ActiveWorkbook.Sheets(sheetName)
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
        ' Adjust asset column dynamically based on sheet name
        Dim assetCol As String
        If sheetName = "Structure Anomalies" Then
            assetCol = "F"
        ElseIf sheetName = "Pipeline Anomalies" Then
            assetCol = "E"
        End If
    
        For i = 2 To lastRow
            asset = NormalizeText(ws.Cells(i, assetCol).Text)
            anomalyRef = NormalizeText(ws.Cells(i, "A").Text)
            If asset <> "" And anomalyRef <> "" Then
                If Not dict.Exists(asset) Then
                    dict(asset) = anomalyRef
                Else
                    dict(asset) = dict(asset) & vbLf & anomalyRef
                End If
            End If
        Next i
        
        ws.Activate
        ws.Cells(1, "A").Select
        ws.Range("A1").AutoFilter
    Next sheetName

    ' Populate Completion Report
    lastRowCompletion = wsCompletion.Cells(wsCompletion.Rows.Count, "E").End(xlUp).Row
    For i = 2 To lastRowCompletion
        If i = 79 Then
            Debug.Print "test"
        End If
    
        asset = NormalizeText(wsCompletion.Cells(i, "E").Value)
        If dict.Exists(asset) Then
            wsCompletion.Cells(i, "R").Value = dict(asset)
        Else
            wsCompletion.Cells(i, "R").Value = dict(asset)
        End If
    Next i
    wsCompletion.Activate
    wsCompletion.Cells(1, "A").Select
    wsCompletion.Range("A1").AutoFilter

    MsgBox "Anomaly References combined from both sheets and populated successfully.", vbInformation
End Sub

Private Sub Find_MissingFiles()
    Dim oFiles As Collection
    Dim oVideos As New Collection
    Dim sFilesFolder As String
    
    Dim iRow As Long
    
    Dim iFolderCol As Long, iFirstFileCol As Long, iLastFileCol As Long, iOrigFirstFile As Long
    Dim iWorkpackCol As Long, iInstallationCol As Long
    
    Dim sFolder As String, sFirstFile As String, sLastValue
    Dim sFileSet As String, sExt As String
    Dim sTemp As String
    
    Dim iTemp As Long
    
    Dim sWorkpack As String, sInstallation As String
    
    ForceFindExtents
    
    iFirstFileCol = FindFirstColumn(Array("First_File", "First File"))
    iOrigFirstFile = Find_Column("Orig First File")
    If iOrigFirstFile = -1 Then iOrigFirstFile = Insert_Column("Orig First File", iFirstFileCol)
    iFirstFileCol = FindFirstColumn(Array("First_File", "First File"))
    iLastFileCol = FindFirstColumn(Array("Last_File", "Last File"))
    
    iFolderCol = Find_Column("Folder")
    iWorkpackCol = Find_Column("Workpack")
    iInstallationCol = Find_Column("Installation")
    
    sFilesFolder = ""
    
    For iRow = 2 To FLastRow
        Cells(iRow, 1).Select
        
        If iFolderCol > 0 Then
            sFolder = Cells(iRow, iFolderCol).Value
        Else
            sInstallation = LCase(Cells(iRow, iInstallationCol).Value)
            sWorkpack = Replace(LCase(Cells(iRow, iWorkpackCol).Value), "/", "_")
            
            sFolder = "Z:\Structures\Routine\" + sInstallation + "\" + sWorkpack + "\"
        End If
        sFirstFile = Cells(iRow, iFirstFileCol).Value
        
        ' Prepare the collection
        If sFilesFolder <> sFolder Then
            Debug.Print ("Preparing list of files in " & sFolder)
            Set oFiles = GetFiles(sFolder)
            Set oFiles = PrepareFileCollection(oFiles)
            sFilesFolder = sFolder
        End If
        
        ' Now we have a sorted list of all files in this folder, lets find the first & last file in the set...
        sFirstFile = Cells(iRow, iFirstFileCol).Value
        
        sTemp = ExtractFilenameOnly(sFirstFile)
        sExt = Trim(ExtractFilenameExt(sFirstFile))
        If sExt = "" Then sExt = ".mp4"
        iTemp = Find_Last(sFirstFile, "_")
        sFileSet = Mid(sTemp, 1, iTemp - 1)
        
        Cells(iRow, iLastFileCol).Select
        Cells(iRow, iOrigFirstFile).Value = sFirstFile
        Cells(iRow, iFirstFileCol).Value = sFileSet & "_0" & sExt ' TODO Actually look for the first file
        Cells(iRow, iLastFileCol).Value = Find_LastFile(oFiles, sFileSet, sExt)
    Next iRow
End Sub

Private Function PrepareFileCollection(AFiles As Collection) As Collection
    Dim iFile As Long
    Dim sFile As String, sExt As String
    Dim sTemp As String
    
    ' Ensures only video files are in the collection
    ' strips path name from the entry, leaving only the filename
    
    For iFile = AFiles.Count To 1 Step -1
        sFile = AFiles(iFile)
        sExt = LCase(ExtractFilenameExt(sFile))
        
        If Not ((sExt = ".mp4") Or (sExt = ".wmv")) Then
            AFiles.Remove (iFile)
        Else
            sTemp = ExtractFilename(sFile)
            
            AFiles.Remove (iFile)
            AFiles.Add (sTemp)
        End If
    Next iFile
    
    ' Set oFiles = CollectionBubbleSort(oFiles)

    Set PrepareFileCollection = AFiles
End Function

Private Function DebugPrint(AInput As Collection)
    Dim element
    
    For Each element In AInput
        Debug.Print element
    Next element
End Function

Private Function Find_LastFile(AFiles As Collection, AFileSet As String, AFileExt As String) As String
    Dim i As Long
    Dim sFile As String
    Dim iCount As Long
    
    ' Right now, this simply counts the number of files in this set,
    ' Last_File is assuming first file starts at 0
    ' and that all other files are contiguous
    ' TODO Confirm this reported entry actually exists in the list
    
    iCount = 0
    
    For i = 1 To AFiles.Count
        sFile = AFiles(i)
        
        If InStr(sFile, AFileSet) > 0 Then iCount = iCount + 1
    Next i
    
    If iCount > 0 Then
        Find_LastFile = AFileSet & "_" & iCount - 1 & AFileExt
    Else
        Find_LastFile = ""
    End If
End Function
