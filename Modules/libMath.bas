Attribute VB_Name = "libMath"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass

Option Explicit

Public Function Math_Max(a As Variant, b As Variant) As Variant
    If VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
        Math_Max = (a Or b) ' Logical max
    ElseIf VarType(a) = vbString And VarType(b) = vbString Then
        If StrComp(a, b, vbBinaryCompare) >= 0 Then
            Math_Max = a
        Else
            Math_Max = b
        End If
    Else
        If a >= b Then
            Math_Max = a
        Else
            Math_Max = b
        End If
    End If
End Function

Public Function Math_Min(a As Variant, b As Variant) As Variant
    If VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
        Math_Min = (a And b) ' Logical min
    ElseIf VarType(a) = vbString And VarType(b) = vbString Then
        If StrComp(a, b, vbBinaryCompare) <= 0 Then
            Math_Min = a
        Else
            Math_Min = b
        End If
    Else
        If a <= b Then
            Math_Min = a
        Else
            Math_Min = b
        End If
    End If
End Function
