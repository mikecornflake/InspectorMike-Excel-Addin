Attribute VB_Name = "libTable"
Option Explicit
Option Private Module

Public Sub BasicTidy(ByVal pWS As Worksheet, Optional ByVal pUseFilter As Boolean = True)
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
    
    If pUseFilter Then
        If pWS.AutoFilterMode Then pWS.AutoFilterMode = False
        If Trim$(CStr(pWS.Cells(1, 1).Value)) <> vbNullString Then
            pWS.Range("A1").AutoFilter
        End If
    End If
    
    Worksheet_FreezeTopRow pWS
    
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
End Sub

Public Sub SortTable(ByVal pWS As Worksheet, Optional ByVal pColumn As Long = -1)
    Dim i As Long
    Dim lLastRow As Long
    Dim lLastColumn As Long
    Dim rngTable As Range

    Call FindTableExtents(pWS, lLastRow, lLastColumn)

    Set rngTable = pWS.Range(pWS.Cells(1, 1), pWS.Cells(lLastRow, lLastColumn))

    If pColumn = -1 Then
        For i = 1 To lLastColumn
            rngTable.Sort _
                Key1:=pWS.Cells(1, i), _
                Order1:=xlAscending, _
                Header:=xlYes
        Next i
    Else
        rngTable.Sort _
            Key1:=pWS.Cells(1, pColumn), _
            Order1:=xlAscending, _
            Header:=xlYes
    End If
End Sub

Public Function CompareRows(ByVal pWS As Worksheet, ByVal pRow1 As Long, ByVal pRow2 As Long, ByVal pLastColumn As Long) As Boolean
    Dim bTemp As Boolean
    Dim i As Long
    
    bTemp = True
    i = 1
    
    While bTemp And (i <= pLastColumn)
        bTemp = (pWS.Cells(pRow1, i).Value = pWS.Cells(pRow2, i).Value)
        
        i = i + 1
    Wend
    
    CompareRows = bTemp
End Function

' pStartRange usually = "A1"
Public Function TableRange(ByVal pWS As Worksheet, Optional ByVal pStartCell As String = "A1", _
                  Optional ByVal pEndRow As Long = -1, Optional ByVal pEndColumn As Long = -1) As Range

    If pEndRow = -1 Or pEndColumn = -1 Then
        Call FindTableExtents(pWS, pEndRow, pEndColumn)
    End If

    Set TableRange = pWS.Range(pWS.Range(pStartCell), pWS.Cells(pEndRow, pEndColumn))

End Function

Public Function DeleteDuplicateRows(ByVal pWS As Worksheet) As Long
    Dim i As Long
    Dim iLastRow As Long, iLastCol As Long
    Dim iDeleted As Long
    
    Call FindTableExtents(pWS, iLastRow, iLastCol)
    Call SortTable(pWS)
    
    iDeleted = 0
    
    For i = iLastRow To 2 Step -1
        If CompareRows(pWS, i, i - 1, iLastCol) Then
            pWS.Rows(i).Delete Shift:=xlUp
            
            iDeleted = iDeleted + 1
        End If
    Next i
    
    DeleteDuplicateRows = iDeleted
End Function

Public Sub DeleteDuplicateRowsByColumn(ByVal pWS As Worksheet, ByVal pColumn As Long)
    Dim i As Long
    
    For i = LastUsedRow(pWS) To 2 Step -1
        If pWS.Cells(i, pColumn).Value = pWS.Cells(i - 1, pColumn).Value Then
            pWS.Rows(i & ":" & i).Delete Shift:=xlUp
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
        FindInColumn = foundCell.row
    End If
End Function

' Return the column number for sName.  if sName doesn't exist, then this column is created
Public Function EnsureColumn(ByVal pWS As Worksheet, ByVal pName As String) As Long
    EnsureColumn = FindColumn(pWS, pName)
    
    If EnsureColumn = -1 Then
        EnsureColumn = AppendColumn(pWS, pName)
    End If
End Function

Public Function DeleteColumn(ByVal pWS As Worksheet, ByVal pName As String) As Boolean
    Dim iCol As Long
    Dim iEndRow As Long
    
    DeleteColumn = False
    iCol = FindColumn(pWS, pName)
    
    If iCol <> -1 Then
        iEndRow = LastUsedRow(pWS)
        
        ' pWS.Columns(iCol).Delete
        pWS.Range(pWS.Cells(1, iCol), pWS.Cells(iEndRow, iCol)).Delete xlShiftToLeft
        
        DeleteColumn = True
    End If
End Function

Public Function RenameColumn(ByVal pWS As Worksheet, ByVal pOldName As String, ByVal pNewName As String) As Boolean
    Dim i As Long
    
    i = FindColumn(pWS, pOldName)
    
    RenameColumn = False
    
    If (i <> -1) Then
        pWS.Cells(1, i).Value = pNewName
        RenameColumn = True
    End If
End Function

' This function needs to be passed an Array
' For use finding columns when the name is subject to minor change (ie "depth" & "depth (m)")
'   iCol = FindFirstColumn(pWS, Array("depth", "depth (m)"))
Public Function FindFirstColumn(ByVal pWS As Worksheet, ByVal pNames) As Long
    Dim i As Long, iCol As Long
    Dim sTemp As String
    
    i = LBound(pNames)
    iCol = -1
    
    While (i <= UBound(pNames)) And (iCol = -1)
        sTemp = pNames(i)
        iCol = FindColumn(pWS, sTemp)
        i = i + 1
    Wend
    
    FindFirstColumn = iCol
End Function

Public Function FindColumn(ByVal pWS As Worksheet, ByVal pHeaderName As String) As Long
    Dim iCol As Long
    Dim sHeader As String

    FindColumn = -1

    For iCol = 1 To LastUsedColumn(pWS)
        sHeader = Trim$(CStr(pWS.Cells(1, iCol).Value))
        If StrComp(sHeader, pHeaderName, vbTextCompare) = 0 Then
            FindColumn = iCol
            Exit Function
        End If
    Next iCol
End Function

Public Function InsertColumn(ByVal pWS As Worksheet, ByVal pName As String, ByVal pCol As Long, Optional ByVal pQuiet As Boolean = True) As Long
    If FindColumn(pWS, pName) <> -1 Then
        If Not pQuiet Then
            MsgBox "Error. Column " & pName & " already exists."
        End If
        
        InsertColumn = -1
    Else
        pWS.Columns(pCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        pWS.Cells(1, pCol).Value = pName
        
        InsertColumn = pCol
    End If
End Function

Public Function AppendColumn(ByVal pWS As Worksheet, ByVal pName As String, Optional ByVal pQuiet As Boolean = True) As Long
    AppendColumn = InsertColumn(pWS, pName, LastUsedColumn(pWS) + 1, pQuiet)
End Function

Public Function GetColumnLetter(ByVal pColumnNumber As Long) As String
    Dim columnLetter As String
    Dim modulo As Long

    columnLetter = ""

    Do
        ' Calculate the modulo (remainder)
        modulo = (pColumnNumber - 1) Mod 26

        ' Convert the modulo to a letter and add it to the column letter
        columnLetter = Chr(65 + modulo) & columnLetter

        ' Calculate the integer division
        pColumnNumber = (pColumnNumber - modulo) \ 26
    Loop While pColumnNumber > 0

    GetColumnLetter = columnLetter
End Function

Public Function IsRowPopulated(ByVal pWS As Worksheet, ByVal pRow As Long) As Boolean
    Dim lastCol As Long
    Dim iCol As Long
    
    IsRowPopulated = False
    
    If pRow < 2 Then Exit Function
    
    lastCol = pWS.Cells(1, pWS.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function
    
    For iCol = 1 To lastCol
        If Len(Trim$(CStr(pWS.Cells(pRow, iCol).Value))) > 0 Then
            IsRowPopulated = True
            Exit Function
        End If
    Next iCol
End Function

Public Sub FormatColumn(ByVal pWS As Worksheet, ByVal pColumn As Long, ByVal pFormat As String)
    If pColumn <> -1 Then
        pWS.Columns(pColumn).NumberFormat = pFormat
    End If
End Sub

Public Sub FormatColumnByName(ByVal pWS As Worksheet, ByVal pName As String, ByVal pFormat As String)
    Dim iCol As Long
    
    iCol = FindColumn(pWS, pName)
    
    Call FormatColumn(pWS, iCol, pFormat)
End Sub

' pNames is an array
Public Sub FormatColumnByNames(ByVal pWS As Worksheet, ByVal pNames, ByVal pFormat As String)
    Dim iCol As Long
    Dim i As Long
    Dim sTemp As String
    
    i = LBound(pNames)
    iCol = -1
    
    While (i <= UBound(pNames))
        sTemp = pNames(i)
        iCol = FindColumn(pWS, sTemp)
        If iCol > 0 Then Call FormatColumn(pWS, iCol, pFormat)
        i = i + 1
    Wend
End Sub

Public Sub ConvertColumnToValues(ByVal pWS As Worksheet, ByVal pColumnName As String, Optional ByVal pCol As Long = -1)
    Dim iCol As Long

    If pCol = -1 Then
        iCol = FindColumn(pWS, pColumnName)
    Else
        iCol = pCol
    End If

    If iCol <> -1 Then
        pWS.Columns(iCol).Value = pWS.Columns(iCol).Value
    End If
End Sub

Public Function LookupColumnValue(ByVal pWS As Worksheet, ByVal pLookupCol As String, ByVal pLookupValue As String, ByVal pReturnCol As String) As String
    Dim iLookupCol As Long
    Dim iReturnCol As Long
    Dim iRow As Long
    Dim sTemp As String

    LookupColumnValue = ""

    iLookupCol = FindColumn(pWS, pLookupCol)
    iReturnCol = FindColumn(pWS, pReturnCol)

    If (iLookupCol = -1) Or (iReturnCol = -1) Then Exit Function

    For iRow = 2 To LastUsedRow(pWS)
        sTemp = Trim$(CStr(pWS.Cells(iRow, iLookupCol).Value))

        If StrComp(sTemp, pLookupValue, vbTextCompare) = 0 Then
            LookupColumnValue = CStr(pWS.Cells(iRow, iReturnCol).Value)
            Exit Function
        End If
    Next iRow
End Function


Public Function PopulateColumn(ByVal pWS As Worksheet, ByVal pColumn As String, ByVal pValue As String) As Boolean
    Dim iCol As Long, iLastRow As Long
    
    iCol = FindColumn(pWS, pColumn)
    iLastRow = LastUsedRow(pWS)
    
    PopulateColumn = False
    
    If iCol <> -1 And iLastRow >= 2 Then
        pWS.Range(pWS.Cells(2, iCol), pWS.Cells(iLastRow, iCol)).Value = pValue
        PopulateColumn = True
    End If
End Function

Public Function MoveColumn(ByVal pWS As Worksheet, ByVal pName As String, ByVal pDestColumn As Long) As Boolean
    Dim iColumn As Long

    MoveColumn = False

    iColumn = FindColumn(pWS, pName)

    If (iColumn <> -1) And (pDestColumn <> -1) And (pDestColumn <> iColumn) Then
        pWS.Columns(iColumn).Cut
        pWS.Columns(pDestColumn).Insert Shift:=xlToRight

        MoveColumn = True
    End If
End Function

Public Function MoveColumnToName(ByVal pWS As Worksheet, ByVal pName As String, ByVal pDestName As String) As Boolean
    Dim iColumn As Long

    iColumn = FindColumn(pWS, pDestName)
    MoveColumnToName = MoveColumn(pWS, pName, iColumn)
End Function

Public Function CopyColumn(ByVal pWS As Worksheet, ByVal pName As String, ByVal pNewName As String) As Boolean
    Dim iCol As Long
    
    iCol = FindColumn(pWS, pName)
    
    CopyColumn = False
    If iCol <> -1 Then
        pWS.Columns(iCol).Copy
        pWS.Columns(iCol + 1).Insert Shift:=xlToRight
        pWS.Cells(1, iCol + 1).Value = pNewName
        
        CopyColumn = True
    End If
End Function

' Helper routine
Public Sub FindTableExtents(ByVal pWS As Worksheet, ByRef pLastRow As Long, ByRef pLastColumn As Long)
    pLastRow = LastUsedRow(pWS)
    pLastColumn = LastUsedColumn(pWS)
End Sub

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
        LastUsedRow = lastCell.row
    End If
End Function

Public Function LastUsedColumn(ByVal pWS As Worksheet) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = pWS.Cells.Find(What:="*", _
                                  After:=pWS.Cells(1, 1), _
                                  LookIn:=xlFormulas, _
                                  LookAt:=xlPart, _
                                  SearchOrder:=xlByColumns, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = lastCell.Column
    End If
End Function
