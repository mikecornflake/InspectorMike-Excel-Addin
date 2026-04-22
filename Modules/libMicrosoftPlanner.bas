Attribute VB_Name = "libMicrosoftPlanner"
Sub ProcessMicrosftPlannerExport()
'
' Macro1 Macro
'

'
    Dim sPlanID, sLink As String
    sPlanID = Cells(2, 2).Value
    
    'Delete the first 4 rows (report style info only, not data)
    Rows("1:4").Select
    Selection.Delete Shift:=xlUp
    
    ForceFindExtents
    
    ' Fix Date Columns
    PlannerFixDateColumnByName ("Created Date")
    PlannerFixDateColumnByName ("Start Date")
    PlannerFixDateColumnByName ("Due Date")
    PlannerFixDateColumnByName ("Completed Date")
    
    Call BasicTidy(ActiveSheet)
    
    ' Re-organise the columns
    Move_Column "Labels", 6
    Move_Column "Description", 7
    Move_Column "Checklist Items", 8
    Move_Column "Completed Checklist Items", 9
    Move_Column "Bucket Name", 2
    
    ' Format Columns
    Columns("B:B").Select
    Selection.WrapText = True
    Columns("C:C").Select
    Selection.WrapText = True
    
    ' Now delete the non-Chevron Buckets
    Dim iRow, iCol As Long
    Dim sTemp As String
    iCol = FindColumn("Bucket Name")
    
    iRow = FLastRow
    While iRow <> 1
        Cells(iRow, 1).Select
        sTemp = Cells(iRow, iCol).Value
        
        If InStr(sTemp, "1003111") = 0 Then
            Rows(iRow).Delete
        End If
        
        iRow = iRow - 1
    Wend
    
    ForceFindExtents
    Cells(2, 1).Select
    
    ' Now sort the data appropriately
    Dim iBucketCol As Long, iProgressCol As Long, iCreatedCol As Long, iCompletedCol As Long, iLabelsCol As Long
    Dim iChecklistCol As Long, iChecklistProgCol As Long
    Dim iComplete As Integer, iTotal As Integer
    
    iBucketCol = FindColumn("Bucket Name")
    iProgressCol = FindColumn("Progress")
    iCreatedCol = FindColumn("Created Date")
    iCompletedCol = FindColumn("Completed Date")
    iLabelsCol = FindColumn("Labels")
    
    iChecklistCol = FindColumn("Checklist Items")
    iChecklistProgCol = FindColumn("Completed Checklist Items")
    
    ' Sort the Rows appropriately
    SelectTable "A1", FLastRow, FLastColumn
    Selection.Sort Key1:=Cells(1, iCreatedCol), Order1:=xlAscending, Header:=xlYes
    Selection.Sort Key1:=Cells(1, iCompletedCol), Order1:=xlAscending, Header:=xlYes
    Selection.Sort Key1:=Cells(1, iProgressCol), Order1:=xlAscending, Header:=xlYes
    Selection.Sort Key1:=Cells(1, iBucketCol), Order1:=xlAscending, Header:=xlYes
    
    ' Show the Checklist Items correctly (split "ChecklistItem1;ChecklistItem2" into lines)
    Columns(iChecklistCol).Select
    Selection.Replace What:=";", Replacement:="" & Chr(10) & "", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    'Colour Code appropriately
    Dim dtMondayLast, dtTemp As Date
    dtMondayLast = DateAdd("ww", -1, Date - (Weekday(Date, vbMonday) - 1))
    
    iRow = 2
    While iRow <= FLastRow
        Cells(iRow, 1).Select
        sTemp = Cells(iRow, 1).Value
        sLink = "https://tasks.office.com/dof.com/en/Home/Planner#/plantaskboard?planId=" & sPlanID & "&taskId=" & sTemp
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=sLink, TextToDisplay:="Task"
        
        Cells(iRow, iProgressCol).Select
        sTemp = Cells(iRow, iProgressCol).Value
        
        ' Completed = Gray Text
        If sTemp = "Completed" Then
            Rows(iRow).Select
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.499984740745262
            End With
            
            dtTemp = Cells(iRow, iCompletedCol).Value
            
            ' Hide items closed before last monday
            If dtTemp < dtMondayLast Then
                Rows(iRow).Select
                Selection.EntireRow.Hidden = True
            End If
        Else
            sTemp = Trim(Cells(iRow, iCreatedCol).Value)
            If sTemp <> "" Then
                dtTemp = Cells(iRow, iCreatedCol).Value
                
                ' New Task = Background Green
                If dtTemp >= dtMondayLast Then
                    Cells(iRow, 2).Select
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent3
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                End If
            End If
        End If
        
        ' Can we do anything easy with the Checklists...
        sTemp = Cells(iRow, iChecklistProgCol).Value
        If sTemp = "" Then
            iComplete = 0
            iTotal = 0
        Else
            iComplete = Val(StringBetween(sTemp, "", "/"))
            iTotal = Val(StringBetween(sTemp, "/", ""))
        End If
        
        Cells(iRow, iChecklistCol).Select
        If (iComplete = iTotal) And (iTotal <> 0) Then
            Selection.Font.Strikethrough = True
        ElseIf (iComplete <> 0) Then
            Selection.Font.Italic = True
        End If
        
        ' On hold = Orange Font
        sTemp = Cells(iRow, iLabelsCol).Value
        If (InStr(sTemp, "Hold") > 0) Or (InStr(sTemp, "Info") > 0) Then
            Rows(iRow).Select
            With Selection.Font
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.249977111117893
            End With
        End If

        iRow = iRow + 1
    Wend
    
    ' Get assorted columns sized correctly
    Columns("A:A").ColumnWidth = 6
    Columns("B:B").ColumnWidth = 35
    Columns("C:C").ColumnWidth = 30
    Columns("D:D").ColumnWidth = 10
    Columns("E:E").ColumnWidth = 10
    Columns("F:F").ColumnWidth = 10
    Columns("G:G").ColumnWidth = 42
    Columns("H:H").ColumnWidth = 42
    Columns("I:I").ColumnWidth = 10
    
    Columns("G:G").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' End neatly
    Cells(2, 1).Select
End Sub

Sub PlannerFixDateColumnByName(AColumn As String)
    Dim iCol As Long
    
    iCol = FindColumn(AColumn)
    
    If iCol <> -1 Then
        PlannerFixDateColumn (iCol)
    End If
End Sub


Sub PlannerFixDateColumn(AColumn As Long)
    ' Microsoft Teams exports dates in "MM/DD/YYYY".
    ' Microsoft Excel doesn't always correctly guess this format, so we're going to have to sort this out manually
    
    
    'Insert the new Calculated Column which will be used to repair the date
    Columns(AColumn + 1).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Build up the Calculated Column
    ' Header
    Cells(1, AColumn + 1).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    
    ' Data
    Cells(2, AColumn + 1).Select
    ActiveCell.FormulaR1C1 = "=IF(TRIM(RC[-1])<>"""", DATE(RIGHT(RC[-1], 4), LEFT(RC[-1], 2), MID(RC[-1], 4, 2)), """")"
    
    ' Now copy this down the table
    Cells(2, AColumn + 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, AColumn + 1), Cells(FLastRow, AColumn + 1))
    
    ' Now "Paste Valves" this Calculated Column over itself
    Columns(AColumn + 1).Select
    Selection.NumberFormat = "dd mmm yyyy"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' And now delete the original
    Columns(AColumn).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
End Sub


