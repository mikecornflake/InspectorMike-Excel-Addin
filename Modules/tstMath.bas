Attribute VB_Name = "tstMath"
Option Explicit
Option Private Module

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
