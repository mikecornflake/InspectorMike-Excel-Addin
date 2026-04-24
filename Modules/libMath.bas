Attribute VB_Name = "libMath"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass

Public Sub Test_LibraryMath()
    ActiveTestModule = "libMath"
    
    ' ===== Basic min/max function tests =====
    Call AssertEqual("Math_Max - 5 vs 10", 10, Math_Max(5, 10))
    Call AssertEqual("Math_Min - 5 vs 10", 5, Math_Min(5, 10))
    Call AssertEqual("Math_Max - negative", -2, Math_Max(-5, -2))
    Call AssertEqual("Math_Min - decimal", 2.5, Math_Min(2.5, 9.8))
    Call AssertEqual("Math_Max - string compare", "zebra", Math_Max("apple", "zebra"))
    Call AssertEqual("Math_Min - string compare", "apple", Math_Min("apple", "zebra"))
    Call AssertEqual("Math_Max - boolean", True, Math_Max(True, False))
    Call AssertEqual("Math_Min - boolean", False, Math_Min(True, False))

    ' ===== Edge cases =====
    Call AssertEqual("Math_Max - empty", "abc", Math_Max("", "abc"))
    Call AssertEqual("min - null vs number", 10, Math_Min(Null, 10))
    Call AssertEqual("Math_Max - error ignored", CVErr(2000), Math_Max(CVErr(2000), CVErr(2000))) ' Custom test, expect error to echo
End Sub

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
