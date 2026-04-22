Attribute VB_Name = "libConditionalFormatting"
' None of this code is generic.  Leaving here as reference for when I do implement a Conditional Formatting library

Public Sub Fix_Planner()
    Clear_Conditional
    
    Add_Conditional_Colour_Weekends
    Add_Conditional_Colour_Today
    'Add_Conditional_Hash_Weekends
    
    Range("A1").Select
End Sub

Public Sub Clear_Conditional()
Attribute Clear_Conditional.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells.FormatConditions.Delete
End Sub

Public Sub Remove_All_Conditional_Formatting()
    Cells.Select
    Cells.FormatConditions.Delete
End Sub

Private Sub Add_Conditional_Colour_Weekends()
    Range("A1").Activate
    Range("3:3,12:12,55:55").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=A$3=""S"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub Add_Conditional_Colour_Today()
Attribute Add_Conditional_Colour_Today.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A1").Activate
    Range("5:11,14:54").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=VALUE(A$4)=TRUNC(NOW())"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub Add_Conditional_Hash_Weekends()
Attribute Add_Conditional_Hash_Weekends.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A1").Activate
    Range("5:11,14:54").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=A$3=""S"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlGray8
        .PatternColorIndex = xlAutomatic
        .ColorIndex = xlAutomatic
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
