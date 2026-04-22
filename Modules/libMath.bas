Attribute VB_Name = "libMath"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass

Public Sub Test_LibraryMath()
    ActiveTestModule = "libMath"

    Dim wsTest As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' === Create temporary sheet ===
    Set wsTest = CreateTestSheet("UnitTestLibraryMath")
    
    On Error GoTo CleanUp
    
    ' ===== Basic min/max function tests =====
    Call AssertEqual("max - 5 vs 10", 10, max(5, 10))
    Call AssertEqual("min - 5 vs 10", 5, min(5, 10))
    Call AssertEqual("max - negative", -2, max(-5, -2))
    Call AssertEqual("min - decimal", 2.5, min(2.5, 9.8))
    Call AssertEqual("max - string compare", "zebra", max("apple", "zebra"))
    Call AssertEqual("min - string compare", "apple", min("apple", "zebra"))
    Call AssertEqual("max - boolean", True, max(True, False))
    Call AssertEqual("min - boolean", False, min(True, False))

    ' ===== Edge cases =====
    Call AssertEqual("max - empty", "abc", max("", "abc"))
    Call AssertEqual("min - null vs number", 10, min(Null, 10))
    Call AssertEqual("max - error ignored", CVErr(2000), max(CVErr(2000), CVErr(2000))) ' Custom test, expect error to echo

    ' ===== Column level tests (requires setup) =====
    ' A1: "ColAlpha", B1: "ColBeta"
    ' === Setup test data ===
    With wsTest
        .Range("A1").Value = "ColAlpha"
        .Range("B1").Value = "ColBeta"
        .Range("A2:A4").Value = Application.Transpose(Array(11, 33, 7))
        .Range("B2:B4").Value = Application.Transpose(Array(-1, 8, 0))
    End With

    ' === Test Column_Max/Min wrappers ===
    Call AssertEqual("Column_Max wrapper - ColAlpha", 33, Column_Max("ColAlpha"))
    Call AssertEqual("Column_Min wrapper - ColAlpha", 7, Column_Min("ColAlpha"))
    Call AssertEqual("Column_Max wrapper - ColBeta", 8, Column_Max("ColBeta"))
    Call AssertEqual("Column_Min wrapper - ColBeta", -1, Column_Min("ColBeta"))

    Call AssertTrue("Column_Max - unknown returns Null", IsNull(Column_Max("Nope")))
    Call AssertTrue("Column_Min - unknown returns Null", IsNull(Column_Min("FakeHeader")))
    
CleanUp:
    ' === Delete temporary sheet ===
    DeleteTestSheet ("UnitTestLibraryMath")
End Sub

Public Function max(a As Variant, b As Variant) As Variant
    If VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
        max = (a Or b) ' Logical max
    ElseIf VarType(a) = vbString And VarType(b) = vbString Then
        If StrComp(a, b, vbBinaryCompare) >= 0 Then
            max = a
        Else
            max = b
        End If
    Else
        If a >= b Then
            max = a
        Else
            max = b
        End If
    End If
End Function

Public Function min(a As Variant, b As Variant) As Variant
    If VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
        min = (a And b) ' Logical min
    ElseIf VarType(a) = vbString And VarType(b) = vbString Then
        If StrComp(a, b, vbBinaryCompare) <= 0 Then
            min = a
        Else
            min = b
        End If
    Else
        If a <= b Then
            min = a
        Else
            min = b
        End If
    End If
End Function

Public Function Column_Max(AColumnName As String) As Variant
    Dim iCol As Long: iCol = Find_Column(AColumnName)
    
    Column_Max = Null
    If iCol <> -1 Then
        Column_Max = Application.WorksheetFunction.max(Columns(iCol))
    End If
End Function

Public Function Column_Min(AColumnName As String) As Variant
    Dim iCol As Long: iCol = Find_Column(AColumnName)
    
    Column_Min = Null
    If iCol <> -1 Then
        Column_Min = Application.WorksheetFunction.min(Columns(iCol))
    End If
End Function
