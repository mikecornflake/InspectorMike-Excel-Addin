Attribute VB_Name = "libQC"
Option Explicit

Private Sub TestCompareSheets()
    Call CompareSheets("MA", "ABU MA", "QC MA")
End Sub

Public Sub CompareSheets(sSheet1 As String, sSheet2 As String, sQCSheet As String)
    Dim oSheet1 As Worksheet, oSheet2 As Worksheet, oQCSheet As Worksheet
    Dim iMaxRow As Long, iMaxColumn As Long
    
    Set oSheet1 = FindSheet(ActiveWorkbook, sSheet1)
    Set oSheet2 = FindSheet(ActiveWorkbook, sSheet2)
    Set oQCSheet = FindSheet(ActiveWorkbook, sQCSheet)
    
    If oQCSheet Is Nothing Then
        Set oQCSheet = AddSheet(ActiveWorkbook, sQCSheet, oSheet2.Index)
    End If
    
    oQCSheet.Select
    Cells.Select
    Selection.ClearContents
    
    ' Find extents and Copy the header
    oSheet2.Select
    
    ForceFindExtents
    iMaxRow = FLastRow
    iMaxColumn = FLastColumn
    
    oSheet1.Select
    
    ForceFindExtents
    iMaxRow = Math_Max(iMaxRow, FLastRow)
    iMaxColumn = Math_Max(iMaxColumn, FLastColumn)
    
    Rows("1:1").Select
    Selection.Copy
    
    oQCSheet.Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    ' Set the QC Formula
    oQCSheet.Select
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=IF('" + sSheet1 + "'!RC<>'" + sSheet2 + "'!RC, ROW('" + sSheet2 + "'!RC), """")"
    
    ' Copy the QC Formula the correct number of columns
    Range("A2").Select
    Selection.Copy
    Range(Cells(2, 1), Cells(iMaxRow, iMaxColumn)).Select
    
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    ' Finalise the sheet
    Range("A2").Select
    With oQCSheet.Tab
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
    End With
    
    Call BasicTidy(ActiveSheet)
End Sub

