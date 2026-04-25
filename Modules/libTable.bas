Attribute VB_Name = "libTable"
Option Explicit
Option Private Module

Public FLastRow As Long
Public FLastColumn As Long

Public Sub BasicTidy(ByVal pWS As Worksheet, Optional ByVal AUseFilter As Boolean = True)
    Dim i As Long
    Dim lastCol As Long
    Dim headerRange As Range
    
    With pWS.Cells
        .VerticalAlignment = xlCenter
        With .Font
            .Name = "Tahoma"
            .size = 10
            .ThemeFont = xlThemeFontNone
        End With
    End With
    
    If AUseFilter Then
        If pWS.AutoFilterMode Then pWS.AutoFilterMode = False
        If Trim$(CStr(pWS.Cells(1, 1).Value)) <> vbNullString Then
            pWS.Range("A1").AutoFilter
        End If
    End If
    
    FreezeTopRow pWS
    
    lastCol = LastUsedColumn(pWS)
    If lastCol > 0 Then
        Set headerRange = pWS.Range(pWS.Cells(1, 1), pWS.Cells(1, lastCol))
        
        headerRange.Font.Bold = True
        
        With headerRange.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
    End If
    
    pWS.Cells.EntireColumn.AutoFit
    pWS.Cells.EntireRow.AutoFit
    
    For i = 1 To lastCol
        If pWS.Columns(i).ColumnWidth > 80 Then
            pWS.Columns(i).ColumnWidth = 80
        End If
    Next i
    
    pWS.Range("A2").Select
End Sub

Public Sub CopyRows(iStartRow As Long, Optional iEndRow As Long = -1)
    If iEndRow = -1 Then
        Rows(iStartRow).Select
    Else
        Range(iStartRow & ":" & iEndRow).Select
    End If
    
    Selection.Copy
End Sub

Public Sub CopyHeader()
    CopyRows (1)
End Sub

Public Sub PasteRow(iRow As Long)
    Cells(iRow, 1).Select
    ActiveSheet.Paste
End Sub

Public Sub PasteHeader()
    PasteRow (1)
End Sub

Public Sub AppendRow()
    PasteRow (FLastRow) + 1
    
    FLastRow = FLastRow + 1
End Sub

Public Sub CopyTable()
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Range("A1").Select
End Sub

Public Sub CopyTableNoHeader()
    ForceFindExtents
    
    SelectTable "A2", FLastRow, FLastColumn
    
    Range("A1").Select
End Sub

Public Sub PasteTable()
    Range("A1").Select
    Selection.End(xlDown).Select
    
    If ActiveCell.Row = 65536 Then
        Range("A2").Select
    Else
        Range("A" & ActiveCell.Row + 1).Select
    End If
    
    ActiveSheet.Paste
    Range("A1").Select
End Sub

Public Sub PasteUnique()
    PasteTable
    DeleteDuplicates
End Sub

Public Sub SortTable(Optional AColumn As Long = -1)
    Dim i As Long
    
    ForceFindExtents
    
    If AColumn = -1 Then
        For i = 1 To FLastColumn
            SelectTable "A1", FLastRow, FLastColumn
            
            Selection.Sort Key1:=Cells(1, i), Order1:=xlAscending, Header:=xlYes
        Next i
    Else
        SelectTable "A1", FLastRow, FLastColumn
        
        Selection.Sort Key1:=Cells(1, AColumn), Order1:=xlAscending, Header:=xlYes
    End If
End Sub

Public Sub DeleteDuplicates()
    Dim i As Long
    Dim iDeleted As Long
    
    ForceFindExtents
    SortTable
    
    iDeleted = 0
    
    For i = FLastRow To 2 Step -1
        If CompareRows(i, i - 1, FLastColumn) Then
            Rows(i & ":" & i).Select
            Selection.Delete Shift:=xlUp
            
            iDeleted = iDeleted + 1
        End If
    Next i
End Sub

Private Function CompareRows(r1 As Long, r2 As Long, iLastColumn As Long) As Boolean
    Dim bTemp As Boolean
    Dim i As Long
    
    bTemp = True
    i = 1
    
    While bTemp And (i <= iLastColumn)
        bTemp = Cells(r1, i).Value = Cells(r2, i).Value
        
        i = i + 1
    Wend
    
    CompareRows = bTemp
End Function

Public Sub ForceFindExtents()
    FLastRow = 0
    FindExtents
End Sub

Public Sub FindExtents()
    If FLastRow = 0 Then
        If Trim(Cells(1, 2).Value) = "" Then
            FLastColumn = 1
        Else
            Range("A1").Select
            Selection.End(xlToRight).Select
            FLastColumn = ActiveCell.Column
        End If
    
        If Trim(Cells(2, 1).Value) = "" Then
            FLastRow = 1
        Else
            Range("A1").Select
            Selection.End(xlDown).Select
            FLastRow = ActiveCell.Row
        End If
    End If
    
    Cells(2, 1).Select
End Sub

Public Sub SelectTable(sStartRange As String, iEndRow As Long, iEndColumn As Long)
    Range(sStartRange).Select
    Range(Selection, Cells(iEndRow, 1)).Select
    Range(Selection, Cells(iEndRow, iEndColumn)).Select
    Selection.Copy
End Sub

Public Sub DeleteDuplicatesByColumnA()
  DeleteDuplicatesByColumn (1)
End Sub

Public Sub DeleteDuplicatesBySelectedColumn()
  DeleteDuplicatesByColumn (ActiveCell.Column)
End Sub

Public Sub DeleteDuplicatesByColumn(iCol As Long)
    Dim i As Long
    
    ForceFindExtents
    
    For i = FLastRow To 2 Step -1
        If Cells(i, iCol).Value = Cells(i - 1, iCol).Value Then
            Rows(i & ":" & i).Select
            Selection.Delete Shift:=xlUp
        End If
    Next i
End Sub

Public Function FindInColumn(ByVal pWS As Worksheet, _
                             ByVal pCol As Long, _
                             ByVal pSearch As String) As Long
    Dim searchRange As Range
    Dim foundCell As Range
    Dim lastRow As Long
    
    lastRow = LastUsedRow(pWS)
    If lastRow < 2 Then
        FindInColumn = 0
        Exit Function
    End If
    
    Set searchRange = pWS.Range(pWS.Cells(2, pCol), pWS.Cells(lastRow, pCol))
    
    Set foundCell = searchRange.Find(What:=pSearch, _
                                     After:=searchRange.Cells(searchRange.Cells.Count), _
                                     LookIn:=xlValues, _
                                     LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, _
                                     MatchCase:=False)
    
    If foundCell Is Nothing Then
        FindInColumn = 0
    Else
        FindInColumn = foundCell.Row
    End If
End Function

' Return the column number for sName.  if sName doesn't exist, then this column is created
' Does not rely on ForceFindExtents
Public Function Ensure_Column(sName As String) As Long
    Ensure_Column = Find_Column(sName)
    
    If Ensure_Column = -1 Then
        Ensure_Column = Add_Column(sName)
    End If
End Function


Public Function Delete_Column(sName As String) As Boolean
    ' Does not rely on ForceFindExtents
    Dim iCol, iEndRow As Long
    Dim bFound As Boolean
    
    Delete_Column = False
    iCol = Find_Column(sName)
    iEndRow = ActiveSheet.UsedRange.Rows.Count
    
    If iCol <> -1 Then
        Range(Cells(1, iCol), Cells(iEndRow, iCol)).Delete xlShiftToLeft
        
        ' Columns(iCol).Delete
        Delete_Column = True
    End If
End Function

Public Function Rename_Column(sOldName As String, sNewName As String) As Boolean
    Dim i As Long
    
    i = Find_Column(sOldName)
    
    Rename_Column = False
    
    If (i <> -1) Then
        Cells(1, i).Value = sNewName
        Rename_Column = True
    End If
End Function

' This function needs to be passed an Array
' For use finding columns when the name is subject to minor change (ie "depth" & "depth (m)")
'   iCol = FindFirstColumn(Array("depth", "depth (m)"))
Public Function FindFirstColumn(ANames) As Long
    Dim i As Long, iCol As Long
    Dim sTemp As String
    
    i = LBound(ANames)
    iCol = -1
    
    While (i <= UBound(ANames)) And (iCol = -1)
        sTemp = ANames(i)
        iCol = Find_Column(sTemp)
        i = i + 1
    Wend
    
    FindFirstColumn = iCol
End Function

Public Function FindColumnInSheet(ByVal pWS As Worksheet, ByVal pHeaderName As String) As Long
    Dim lastCol As Long
    Dim iCol As Long
    Dim sHeader As String

    FindColumnInSheet = -1

    lastCol = pWS.Cells(1, pWS.Columns.Count).End(xlToLeft).Column

    For iCol = 1 To lastCol
        sHeader = Trim$(CStr(pWS.Cells(1, iCol).Value))
        If StrComp(sHeader, pHeaderName, vbTextCompare) = 0 Then
            FindColumnInSheet = iCol
            Exit Function
        End If
    Next iCol
End Function

Public Function Find_Column(pHeaderName As String) As Long
    Find_Column = FindColumnInSheet(ActiveSheet, pHeaderName)
End Function

' Deprecated
Public Function FindColumn(pHeaderName As String) As Long
    FindColumn = FindColumnInSheet(ActiveSheet, pHeaderName)
End Function

Public Function Insert_Column(AName As String, ACol As Long, Optional bQuiet As Boolean = True) As Long
    ' Does not rely on ForceFindExtents
    If Find_Column(AName) <> -1 Then
        If Not bQuiet Then
            MsgBox "Error.  Column " & AName & " already exists "
        End If
        Insert_Column = -1
    Else
        Columns(ACol).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        Insert_Column = ACol
        Cells(1, Insert_Column).Value = AName
    End If
    
    ' Although we don't rely on extents, we should ensure they're correct if used
    If FLastColumn <> -1 Then
        FLastColumn = FLastColumn + 1
    End If
End Function

Public Function Add_Column(sName As String, Optional bQuiet As Boolean = True) As Long
    ' Does not rely on ForceFindExtents
    Dim iLastColumn As Long
    
    If Find_Column(sName) <> -1 Then
        If Not bQuiet Then
            MsgBox "Error.  Column " & sName & " already exists "
        End If
        Add_Column = -1
    Else
        If Cells(1, 1).Value = "" Then
            Add_Column = 1
        ElseIf Cells(1, 2).Value = "" Then
            Add_Column = 2
        Else
            Range("A1").Select
            
            Selection.End(xlToRight).Select
            iLastColumn = ActiveCell.Column
            
            Add_Column = iLastColumn + 1
        End If
        Cells(1, Add_Column).Value = sName
    End If
    
    ' Although we don't rely on extents, we should ensure they're correct if used
    If FLastColumn <> -1 Then
        FLastColumn = FLastColumn + 1
    End If
End Function

Public Function Copy_Column(sName As String, sNewName As String) As Boolean
    Dim iCol As Long
    
    iCol = Find_Column(sName)
    
    Copy_Column = False
    If iCol <> -1 Then
        Columns(iCol).Select
        Selection.Copy
        Columns(iCol + 1).Select
        Selection.Insert Shift:=xlToRight
        Cells(1, iCol + 1).Value = sNewName
        
        Copy_Column = True
    End If
End Function

Function GetColumnLetter(columnNumber As Long) As String
    Dim dividend As Long
    Dim columnLetter As String
    Dim modulo As Long

    columnLetter = ""

    Do
        ' Calculate the modulo (remainder)
        modulo = (columnNumber - 1) Mod 26

        ' Convert the modulo to a letter and add it to the column letter
        columnLetter = Chr(65 + modulo) & columnLetter

        ' Calculate the integer division
        columnNumber = (columnNumber - modulo) \ 26
    Loop While columnNumber > 0

    GetColumnLetter = columnLetter
End Function

Public Function IsWorksheetRowPopulated(ByVal pWS As Worksheet, ByVal pRow As Long) As Boolean
    Dim lastCol As Long
    Dim iCol As Long
    
    IsWorksheetRowPopulated = False
    
    If pRow < 2 Then Exit Function
    
    lastCol = pWS.Cells(1, pWS.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function
    
    For iCol = 1 To lastCol
        If Len(Trim$(CStr(pWS.Cells(pRow, iCol).Value))) > 0 Then
            IsWorksheetRowPopulated = True
            Exit Function
        End If
    Next iCol
End Function

Public Sub FormatColumn(AColumn As Long, AFormat As String)
    If AColumn <> -1 Then
        Columns(AColumn).Select
        Selection.NumberFormat = AFormat
        'Selection.HorizontalAlignment = xlRight
    End If
End Sub

Public Sub FormatColumnByName(AName As String, AFormat As String)
    Dim iCol As Long
    
    iCol = FindColumn(AName)
    Call FormatColumn(iCol, AFormat)
End Sub

Public Sub FormatColumnByNames(ANames, AFormat As String)
    Dim iCol As Long
    Dim i As Long
    Dim sTemp As String
    
    i = LBound(ANames)
    iCol = -1
    
    While (i <= UBound(ANames))
        sTemp = ANames(i)
        iCol = FindColumn(sTemp)
        If iCol > 0 Then Call FormatColumn(iCol, AFormat)
        i = i + 1
    Wend
End Sub

Public Sub ConvertColumnToValues(AColumnName As String, Optional ACol As Long = -1)
    Dim iCol As Long
    
    If ACol = -1 Then
        iCol = Find_Column(AColumnName)
    Else
        iCol = ACol
    End If
    
    If iCol <> -1 Then
        Columns(iCol).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
End Sub

' Doesn't rely on ForceFindExtents
Public Function Lookup(sLookupCol As String, sLookupValue As String, sReturnCol As String) As String
    Dim iLookupCol As Long
    Dim iReturnCol As Long
    Dim iRow As Long
    Dim bFound As Boolean
    Dim sTemp As String
    
    iLookupCol = Find_Column(sLookupCol)
    iReturnCol = Find_Column(sReturnCol)
    
    iRow = 2
    sTemp = Cells(iRow, iLookupCol).Value
    
    While (Not bFound) And (sTemp <> "")
        iRow = iRow + 1
        sTemp = Cells(iRow, iLookupCol).Value
        bFound = (sTemp = sLookupValue)
    Wend
    
    If bFound Then
        Lookup = Cells(iRow, iReturnCol).Value
    Else
        Lookup = ""
    End If
End Function


Public Function PopulateColumn(sColumn As String, sValue As String) As Boolean
    ' Requires ForceFindExtents
    Dim iCol As Long
    
    iCol = FindColumn(sColumn)
    
    PopulateColumn = False
    
    If iCol <> -1 Then
        Cells(2, iCol).Value = sValue
        Cells(2, iCol).Copy
        Range(Cells(2, iCol), Cells(FLastRow, iCol)).Select
        ActiveSheet.Paste
    
        PopulateColumn = True
    End If
End Function

Public Function Move_Column(sName As String, iNewColumn As Long) As Boolean
    Dim iColumn As Long
    
    Move_Column = False
    iColumn = Find_Column(sName)
    If iColumn <> -1 And iNewColumn <> -1 And iNewColumn <> iColumn Then
        Columns(iColumn).Select
        Selection.Cut
        
        Columns(iNewColumn).Select
        Selection.Insert Shift:=xlToRight
        
        Move_Column = True
    End If
End Function

Public Function Move_Column2(AColToMove As String, ABeforeCol As String) As Boolean
    Dim iColBefore As Long
    
    iColBefore = Find_Column(ABeforeCol)
    Move_Column2 = Move_Column(AColToMove, iColBefore)
End Function
Public Function First_Non_Header_Row() As Long
    First_Non_Header_Row = 2
    
    While Cells(First_Non_Header_Row, 1).Value = ""
        First_Non_Header_Row = First_Non_Header_Row + 1
    Wend
End Function

Public Function LastUsedRow(ByVal pWS As Worksheet) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = pWS.Cells.Find(What:="*", _
                                  After:=pWS.Cells(1, 1), _
                                  LookIn:=xlFormulas, _
                                  LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = lastCell.Row
    End If
End Function

Public Function LastUsedColumn(ByVal pWS As Worksheet) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = pWS.Cells.Find(What:="*", _
                                  After:=pWS.Cells(1, 1), _
                                  LookIn:=xlFormulas, _
                                  LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = lastCell.Column
    End If
End Function

Public Sub Copy_Column_To_Sheet(ASource As Worksheet, ASourceColName As String, ADest As Worksheet, ADestColName As String)
    ASource.Select
    Dim iSourceCol As Long: iSourceCol = Find_Column(ASourceColName)
    
    ADest.Select
    Dim iDestCol As Long: iDestCol = Find_Column(ADestColName)
    
    If (iSourceCol = -1) Or (iDestCol = -1) Then
        MsgBox "Either source " & ASourceColName & " or destination " & ADestColName & " do not exist"
        Exit Sub
    End If
    
    ASource.Select
    Columns(iSourceCol).Select
    Selection.Copy
    
    ADest.Select
    Cells(1, iDestCol).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' The above operation would have overwritten the original Destination name, let's reset it
    Cells(1, iDestCol).FormulaR1C1 = ADestColName
    Cells(1, iDestCol).Select
End Sub
