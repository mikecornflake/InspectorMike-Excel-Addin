Attribute VB_Name = "LibraryTable"
Option Explicit

Public FLastRow As Long
Public FLastColumn As Long

' Assumes header in Row 1 and Col 1 is fully populated
' Turns this into a decent looking and functional table
' complete with decently sized columns, nice font,
' first row frozen, filters on etc
Public Sub BasicTidy(Optional AUseFilter As Boolean = True)
    Dim i As Long
    
    ' Get the general formatting correct
    Cells.Select

    Selection.VerticalAlignment = xlCenter
    With Selection.Font
        .Name = "Tahoma"
        .size = 10
   '     .ColorIndex = xlAutomatic
        .ThemeFont = xlThemeFontNone
    End With
    
    If AUseFilter Then
        ' Turn off AutoFilter if it's already on as it's scope may need widening or moving....
        If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
        
        ' Turn AutoFilter on
        Range("A1").Select
        If Trim(Cells(1, 1).Value) <> "" Then Selection.AutoFilter
    End If
    
    ' Freeze the first row
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    ' Colour the first row
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    
    ' Kerry Preferences
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    ' Mike Preferences
    'With Selection.Interior
    '    .Pattern = xlSolid
    '    .PatternColorIndex = xlAutomatic
    '    .ThemeColor = xlThemeColorLight2
    '    .TintAndShade = 0.799981688894314
    '    .PatternTintAndShade = 0
    'End With
    
    ' Neaten the table
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
    ForceFindExtents
    
    For i = 1 To FLastColumn
        If Columns(i).ColumnWidth > 80 Then
            Columns(i).ColumnWidth = 80
        End If
    Next i
    
    ' Return a sensible selection
    Range("A2").Select
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

Public Sub ApplyFormatting()
    ' Sort by the last column (assumes last column = unique ID)
    Cells(1, 1).Select
    Selection.End(xlToRight).Select
    ActiveCell.Sort Key1:=ActiveCell, Order1:=xlAscending, Header:=xlYes
    
    ' Make the header row bold
    Rows("1:1").Select
    Selection.Font.Bold = True
    
    ' Stretch each column so it fits
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
    ' Freeze the first row
    ActiveWindow.SplitRow = 0.764705882352941
    ActiveWindow.FreezePanes = True
    
    ' move the cursor to the top left, all neat and tidy
    Cells(1, 1).Select
End Sub

Public Sub ApplyFormattingNoSort()
    ' Make the header row bold
    Rows("1:1").Select
    Selection.Font.Bold = True
    
    ' Stretch each column so it fits
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
    ' Freeze the first row
    ActiveWindow.SplitRow = 0.764705882352941
    ActiveWindow.FreezePanes = True
    
    ' move the cursor to the top left, all neat and tidy
    Cells(1, 1).Select
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

Public Sub DeleteDuplicatesByColumn(iCol As Integer)
    Dim i As Long
    
    ForceFindExtents
    
    For i = FLastRow To 2 Step -1
        If Cells(i, iCol).Value = Cells(i - 1, iCol).Value Then
            Rows(i & ":" & i).Select
            Selection.Delete Shift:=xlUp
        End If
    Next i
End Sub

Public Function Find_In_Column(iCol As Long, sSearch As String, Optional ASheet As Worksheet = Nothing) As Long
    Dim i As Long
    Dim bFound As Boolean
    Dim oSheet As Worksheet
    
    If ASheet Is Nothing Then
        Set oSheet = ActiveWorkbook.ActiveSheet
    Else
        Set oSheet = ASheet
    End If
    
    i = 2
    bFound = False
        
    While Not bFound And i <= oSheet.UsedRange.Rows.Count
        bFound = UCase(sSearch) = UCase(oSheet.Cells(i, iCol).Value)
        
        If Not bFound Then
            i = i + 1
        End If
    Wend
    
    If bFound Then
        Find_In_Column = i
    Else
        Find_In_Column = 0
    End If
End Function

' Return the column number for sName.  if sName doesn't exist, then this column is created
' Does not rely on ForceFindExtents
Public Function Ensure_Column(sName As String) As Integer
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
    Dim i As Integer
    
    i = Find_Column(sOldName)
    
    Rename_Column = False
    
    If (i <> -1) Then
        Cells(1, i).Value = sNewName
        Rename_Column = True
    End If
End Function

' Needs to be passed an Array.
' For use finding columns when the name is subject to minor change (ie "depth" & "depth (m)")
'   iCol = FindFirstColumn(Array("depth", "depth (m)"))
Public Function FindFirstColumn(ANames) As Integer
    Dim i As Long, iCol As Integer
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

Public Function FindColumn(sName As String) As Integer
    ' Same as Find_Column, but uses the one-off
    '  ForceFindExtents to determine number of columns
    Dim i As Long
    i = 1
    
    While (i <= FLastColumn) And (UCase(Cells(1, i).Value) <> UCase(sName))
        i = i + 1
    Wend
    
    If i > FLastColumn Then
        FindColumn = -1
    Else
        FindColumn = i
    End If
End Function

Public Function Insert_Column(AName As String, ACol As Long, Optional bQuiet As Boolean = True) As Integer
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

Public Function Add_Column(sName As String, Optional bQuiet As Boolean = True) As Integer
    ' Does not rely on ForceFindExtents
    Dim iLastColumn As Integer
    
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

Function GetColumnLetter(columnNumber As Integer) As String
    Dim dividend As Integer
    Dim columnLetter As String
    Dim modulo As Integer

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

' Only formats the first column found in the array, NOT all of them
Public Sub FormatColumnByNames(ANames, AFormat As String)
    Dim iCol As Long
    Dim i As Long
    Dim sTemp As String
    
    i = LBound(ANames)
    iCol = -1
    
    While (i <= UBound(ANames)) And (iCol = -1)
        sTemp = ANames(i)
        iCol = FindColumn(sTemp)
        i = i + 1
    Wend
    
    Call FormatColumn(iCol, AFormat)
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

Public Function Find_Column(sName As String) As Integer
    ' Doesn't rely on ForceFindExtents
    Dim i As Integer
    Dim iLastColumn As Integer
    Dim bFound As Boolean
    
    ' Find Last Column
    Range("A1").Select
    Selection.End(xlToRight).Select
    iLastColumn = ActiveCell.Column
    
    bFound = False
    i = 1
    
    While Not bFound And (i <= iLastColumn)
        bFound = Trim(UCase(Cells(1, i).Value)) = Trim(UCase(sName))
        
        i = i + 1
    Wend
    
    If bFound Then
        Find_Column = i - 1
    Else
        Find_Column = -1
    End If
End Function

Public Function Move_Column(sName As String, iNewColumn As Integer) As Boolean
    Dim iColumn As Integer
    
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
    Dim iColBefore As Integer
    
    iColBefore = Find_Column(ABeforeCol)
    Move_Column2 = Move_Column(AColToMove, iColBefore)
End Function
Public Function First_Non_Header_Row() As Integer
    First_Non_Header_Row = 2
    
    While Cells(First_Non_Header_Row, 1).Value = ""
        First_Non_Header_Row = First_Non_Header_Row + 1
    Wend
End Function

Public Function Quick_Last_Row() As Integer
    Range("A" & First_Non_Header_Row).Select
    Selection.End(xlDown).Select
    Last_Row = ActiveCell.Row
End Function

Public Function Quick_Last_Column() As Integer
    Range("A1").Select
    Selection.End(xlToRight).Select
    Last_Column = ActiveCell.Column
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
