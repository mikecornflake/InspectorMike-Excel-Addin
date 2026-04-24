Attribute VB_Name = "tstString"
Option Private Module

Private Sub Test_Text_ToBool_Invalid(ByVal pInput As String, ByVal pTestName As String)
    On Error Resume Next
    Err.Clear
    
    Call Text_ToBool(pInput)
    
    Call AssertTrue(pTestName, Err.Number <> 0)
    
    On Error GoTo 0
End Sub

Public Sub Test_LibraryString()
    ActiveTestModule = "libString"

    ' =========================
    ' Tests for Text_ToBool
    ' =========================
    Call AssertTrue("Text_ToBool() - true", Text_ToBool("true"))
    Call AssertTrue("Text_ToBool() - yes", Text_ToBool("yes"))
    Call AssertTrue("Text_ToBool() - TRUE (case)", Text_ToBool("TRUE"))
    Call AssertTrue("Text_ToBool() - y", Text_ToBool("y"))
    Call AssertTrue("Text_ToBool() - 1", Text_ToBool("1"))
    Call AssertTrue("Text_ToBool() - padded true", Text_ToBool("  true  "))
    Call AssertTrue("Text_ToBool() - padded yes", Text_ToBool("  yes  "))
    
    Call AssertFalse("Text_ToBool() - false", Text_ToBool("false"))
    Call AssertFalse("Text_ToBool() - no", Text_ToBool("no"))
    Call AssertFalse("Text_ToBool() - n", Text_ToBool("n"))
    Call AssertFalse("Text_ToBool() - 0", Text_ToBool("0"))
    Call AssertFalse("Text_ToBool() - padded false", Text_ToBool("  false  "))
    Call AssertFalse("Text_ToBool() - padded no", Text_ToBool("  no  "))
    
    ' Invalid inputs should raise errors
    Call Test_Text_ToBool_Invalid("", "Text_ToBool() - empty")
    Call Test_Text_ToBool_Invalid("maybe", "Text_ToBool() - maybe")
    Call Test_Text_ToBool_Invalid("yes!", "Text_ToBool() - yes!")
    Call Test_Text_ToBool_Invalid("2", "Text_ToBool() - 2")
    Call Test_Text_ToBool_Invalid("abc", "Text_ToBool() - abc")
    
    ' =========================
    ' Tests for Bool_ToText
    ' =========================
    Call AssertEqual("Bool_ToText() - True", "True", Bool_ToText(True))
    Call AssertEqual("Bool_ToText() - False", "False", Bool_ToText(False))
    
    
    ' =========================
    ' Tests for Text_IsBool
    ' =========================
    Call AssertTrue("Text_IsBool - true", Text_IsBool("true"))
    Call AssertTrue("Text_IsBool - TRUE", Text_IsBool("TRUE"))
    Call AssertTrue("Text_IsBool - yes", Text_IsBool("yes"))
    Call AssertTrue("Text_IsBool - y", Text_IsBool("y"))
    Call AssertTrue("Text_IsBool - 1", Text_IsBool("1"))
    
    Call AssertTrue("Text_IsBool - false", Text_IsBool("false"))
    Call AssertTrue("Text_IsBool - FALSE", Text_IsBool("FALSE"))
    Call AssertTrue("Text_IsBool - no", Text_IsBool("no"))
    Call AssertTrue("Text_IsBool - n", Text_IsBool("n"))
    Call AssertTrue("Text_IsBool - 0", Text_IsBool("0"))
    
    Call AssertTrue("Text_IsBool - padded true", Text_IsBool("  true  "))
    Call AssertTrue("Text_IsBool - padded false", Text_IsBool("  false  "))
    
    Call AssertFalse("Text_IsBool - empty", Text_IsBool(""))
    Call AssertFalse("Text_IsBool - spaces", Text_IsBool("   "))
    Call AssertFalse("Text_IsBool - Yes!", Text_IsBool("Yes!"))
    Call AssertFalse("Text_IsBool - maybe", Text_IsBool("maybe"))
    Call AssertFalse("Text_IsBool - 2", Text_IsBool("2"))
    Call AssertFalse("Text_IsBool - abc", Text_IsBool("abc"))

    ' Text_Remove
    Call AssertEqual("Text_Remove - middle", "abcxyz", Text_Remove("abc123xyz", "123"))
    Call AssertEqual("Text_Remove - not found", "abc", Text_Remove("abc", "zzz"))

    ' Text_IsNumber
    Call AssertTrue("Text_IsNumber - integer", Text_IsNumber("42"))
    Call AssertTrue("Text_IsNumber - decimal", Text_IsNumber("3.14"))
    Call AssertFalse("Text_IsNumber - text", Text_IsNumber("forty-two"))
    Call AssertFalse("Text_IsNumber - blank", Text_IsNumber(""))

    ' Text_Between
    Call AssertEqual("Text_Between - normal", "123", Text_Between("abc[123]xyz", "[", "]"))
    Call AssertEqual("Text_Between - reverse", "final", Text_Between("start <mid> end <final>", "<", ">", True))
    Call AssertEqual("Text_Between - missing", "", Text_Between("abc", "[", "]"))
    Call AssertEqual("Text_Between - start to delimiter", "Hello", Text_Between("Hello world!", "", " "))
    Call AssertEqual("Text_Between - delimiter to end", "world!", Text_Between("Hello world!", " ", ""))
    Call AssertEqual("Text_Between - full string", "Hello world!", Text_Between("Hello world!", "", ""))
    
    ' Text_Replace
    Call AssertEqual("Text_Replace - found", "abc456xyz", Text_Replace("abc123xyz", "123", "456"))
    Call AssertEqual("Text_Replace - not found", "abc", Text_Replace("abc", "zzz", "xxx"))
    
    ' Text_FindLast
    Call AssertEqual("Text_FindLast - single", 4, Text_FindLast("abc123xyz", "123"))
    Call AssertEqual("Text_FindLast - multiple", 7, Text_FindLast("a-b-c-b", "b"))
    Call AssertEqual("Text_FindLast - not found", 0, Text_FindLast("abc", "z"))
    
    ' Text_AfterLast
    Call AssertEqual("Text_AfterLast - found", "d", Text_AfterLast("a.b.c.d", "."))
    Call AssertEqual("Text_AfterLast - not found", "", Text_AfterLast("abcd", ","))
    
    ' Text_BeforeLast
    Call AssertEqual("Text_BeforeLast - found", "a.b.c", Text_BeforeLast("a.b.c.d", "."))
    Call AssertEqual("Text_BeforeLast - not found", "abcd", Text_BeforeLast("abcd", ","))
    
    ' Text_ToSentenceCase
    Call AssertEqual("Text_ToSentenceCase - basic", "Hello. How are you? I'm fine!", Text_ToSentenceCase("hello. how are you? i'm fine!"))

    ' Text_IsLatin
    Call AssertTrue("Text_IsLatin - basic", Text_IsLatin("Hello µ"))
    Call AssertFalse("Text_IsLatin - non-latin", Text_IsLatin(ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09))) ' Chinese characters: ???
End Sub

