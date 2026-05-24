Sub FormatIEWorkpackDataSheet()
'
' FormatIEWorkpackDataSheet Macro
'

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

Sub SortSheetsAlphabetically()
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