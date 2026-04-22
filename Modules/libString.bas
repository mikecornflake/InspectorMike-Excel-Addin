Attribute VB_Name = "libString"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass

Option Explicit

Private Const ERR_BASE_LIBRARY_STRING As Long = vbObjectError + 1024
Private Const ERR_INVALID_STRING_BOOL As Long = ERR_BASE_LIBRARY_STRING + 1

Private Sub Test_StringBool_Invalid(ByVal pInput As String, ByVal pTestName As String)
    On Error Resume Next
    Err.Clear
    
    Call StringBool(pInput)
    
    Call AssertTrue(pTestName, Err.Number <> 0)
    
    On Error GoTo 0
End Sub

Public Sub Test_LibraryString()
    ActiveTestModule = "libString"

    ' =========================
    ' Tests for StringBool
    ' =========================
    Call AssertTrue("StringBool() - true", StringBool("true"))
    Call AssertTrue("StringBool() - yes", StringBool("yes"))
    Call AssertTrue("StringBool() - TRUE (case)", StringBool("TRUE"))
    Call AssertTrue("StringBool() - y", StringBool("y"))
    Call AssertTrue("StringBool() - 1", StringBool("1"))
    Call AssertTrue("StringBool() - padded true", StringBool("  true  "))
    Call AssertTrue("StringBool() - padded yes", StringBool("  yes  "))
    
    Call AssertFalse("StringBool() - false", StringBool("false"))
    Call AssertFalse("StringBool() - no", StringBool("no"))
    Call AssertFalse("StringBool() - n", StringBool("n"))
    Call AssertFalse("StringBool() - 0", StringBool("0"))
    Call AssertFalse("StringBool() - padded false", StringBool("  false  "))
    Call AssertFalse("StringBool() - padded no", StringBool("  no  "))
    
    ' Invalid inputs should raise errors
    Call Test_StringBool_Invalid("", "StringBool() - empty")
    Call Test_StringBool_Invalid("maybe", "StringBool() - maybe")
    Call Test_StringBool_Invalid("yes!", "StringBool() - yes!")
    Call Test_StringBool_Invalid("2", "StringBool() - 2")
    Call Test_StringBool_Invalid("abc", "StringBool() - abc")
    
    ' =========================
    ' Tests for BoolString
    ' =========================
    Call AssertEqual("BoolString() - True", "True", BoolString(True))
    Call AssertEqual("BoolString() - False", "False", BoolString(False))
    
    
    ' =========================
    ' Tests for ValidStringBool
    ' =========================
    Call AssertTrue("ValidStringBool - true", ValidStringBool("true"))
    Call AssertTrue("ValidStringBool - TRUE", ValidStringBool("TRUE"))
    Call AssertTrue("ValidStringBool - yes", ValidStringBool("yes"))
    Call AssertTrue("ValidStringBool - y", ValidStringBool("y"))
    Call AssertTrue("ValidStringBool - 1", ValidStringBool("1"))
    
    Call AssertTrue("ValidStringBool - false", ValidStringBool("false"))
    Call AssertTrue("ValidStringBool - FALSE", ValidStringBool("FALSE"))
    Call AssertTrue("ValidStringBool - no", ValidStringBool("no"))
    Call AssertTrue("ValidStringBool - n", ValidStringBool("n"))
    Call AssertTrue("ValidStringBool - 0", ValidStringBool("0"))
    
    Call AssertTrue("ValidStringBool - padded true", ValidStringBool("  true  "))
    Call AssertTrue("ValidStringBool - padded false", ValidStringBool("  false  "))
    
    Call AssertFalse("ValidStringBool - empty", ValidStringBool(""))
    Call AssertFalse("ValidStringBool - spaces", ValidStringBool("   "))
    Call AssertFalse("ValidStringBool - Yes!", ValidStringBool("Yes!"))
    Call AssertFalse("ValidStringBool - maybe", ValidStringBool("maybe"))
    Call AssertFalse("ValidStringBool - 2", ValidStringBool("2"))
    Call AssertFalse("ValidStringBool - abc", ValidStringBool("abc"))

    ' Compare
    Call AssertTrue("Compare - match", Compare("Hello", " hello "))
    Call AssertFalse("Compare - no match", Compare("Hello", "world"))

    ' RemoveSubString
    Call AssertEqual("RemoveSubString - middle", "abcxyz", RemoveSubString("abc123xyz", "123"))
    Call AssertEqual("RemoveSubString - not found", "abc", RemoveSubString("abc", "zzz"))

    ' IsNumber
    Call AssertTrue("IsNumber - integer", IsNumber("42"))
    Call AssertTrue("IsNumber - decimal", IsNumber("3.14"))
    Call AssertFalse("IsNumber - text", IsNumber("forty-two"))
    Call AssertFalse("IsNumber - blank", IsNumber(""))

    ' StringBetween
    Call AssertEqual("StringBetween - normal", "123", StringBetween("abc[123]xyz", "[", "]"))
    Call AssertEqual("StringBetween - reverse", "final", StringBetween("start <mid> end <final>", "<", ">", True))
    Call AssertEqual("StringBetween - missing", "", StringBetween("abc", "[", "]"))
    Call AssertEqual("StringBetween - start to delimiter", "Hello", StringBetween("Hello world!", "", " "))
    Call AssertEqual("StringBetween - delimiter to end", "world!", StringBetween("Hello world!", " ", ""))
    Call AssertEqual("StringBetween - full string", "Hello world!", StringBetween("Hello world!", "", ""))
    
    ' SwapString
    Call AssertEqual("SwapString - found", "abc456xyz", SwapString("abc123xyz", "123", "456"))
    Call AssertEqual("SwapString - not found", "abc", SwapString("abc", "zzz", "xxx"))
    
    ' Find_Last
    Call AssertEqual("Find_Last - single", 4, Find_Last("abc123xyz", "123"))
    Call AssertEqual("Find_Last - multiple", 7, Find_Last("a-b-c-b", "b"))
    Call AssertEqual("Find_Last - not found", 0, Find_Last("abc", "z"))
    
    ' StringAfterLast
    Call AssertEqual("StringAfterLast - found", "d", StringAfterLast("a.b.c.d", "."))
    Call AssertEqual("StringAfterLast - not found", "", StringAfterLast("abcd", ","))
    
    ' StringBeforeLast
    Call AssertEqual("StringBeforeLast - found", "a.b.c", StringBeforeLast("a.b.c.d", "."))
    Call AssertEqual("StringBeforeLast - not found", "abcd", StringBeforeLast("abcd", ","))
    
    ' SentenceCase
    Call AssertEqual("SentenceCase - basic", "Hello. How are you? I'm fine!", SentenceCase("hello. how are you? i'm fine!"))

    ' IsLatin
    Call AssertTrue("IsLatin - basic", IsLatin("Hello µ"))
    Call AssertFalse("IsLatin - non-latin", IsLatin(ChrW(&H4E00) & ChrW(&H4E8C) & ChrW(&H4E09))) ' Chinese characters: ???
End Sub

Public Function StringBool(ByVal AInput As String) As Boolean
    Dim sInput As String
    
    sInput = LCase$(Trim$(AInput))
    
    Select Case sInput
        Case "true", "yes", "y", "1"
            StringBool = True
        
        Case "false", "no", "n", "0"
            StringBool = False
        
        Case Else
            Err.Raise _
                Number:=ERR_INVALID_STRING_BOOL, _
                source:="LibraryString.StringBool", _
                Description:="Invalid boolean string: [" & AInput & "]"
    End Select
End Function

Public Function BoolString(ByVal AInput As Boolean) As String
    If AInput Then
        BoolString = "True"
    Else
        BoolString = "False"
    End If
End Function

Public Function ValidStringBool(ByVal AInput As String) As Boolean
    Dim sInput As String
    
    sInput = LCase$(Trim$(AInput))
    
    Select Case sInput
        Case "true", "yes", "y", "1", _
             "false", "no", "n", "0"
            ValidStringBool = True
        
        Case Else
            ValidStringBool = False
    End Select
End Function

Public Function Compare(sString1 As String, sString2 As String) As Boolean
    Compare = Trim(LCase(sString1)) = Trim(LCase(sString2))
End Function

Public Function RemoveSubString(sInput As String, sSubString As String) As String
    Dim iSub As Integer
    Dim iLen As Integer
    
    iSub = InStr(sInput, sSubString)
    iLen = Len(sSubString)
    
    If iSub <> 0 Then
        RemoveSubString = Left(sInput, iSub - 1) & Mid(sInput, iSub + iLen)
    Else
        RemoveSubString = sInput
    End If
End Function

Public Function IsNumber(sValue As String) As Boolean
    On Error GoTo ErrorHandler
    
    sValue = Trim(sValue)
    
    IsNumber = ("" & Val(sValue)) = sValue
    
    Exit Function
ErrorHandler:
    On Error Resume Next
    IsNumber = False
End Function

' Developer: copilot.  15/08/2025
Public Function StringBetween(strMain As String, str1 As String, str2 As String, Optional reverse As Boolean = False) As String
    Dim startPos As Long, endPos As Long

    ' Handle start delimiter
    If str1 = "" Then
        startPos = 1
    ElseIf reverse Then
        startPos = InStrRev(strMain, str1)
        If startPos = 0 Then Exit Function
        startPos = startPos + Len(str1)
    Else
        startPos = InStr(strMain, str1)
        If startPos = 0 Then Exit Function
        startPos = startPos + Len(str1)
    End If

    ' Handle end delimiter
    If str2 = "" Then
        endPos = Len(strMain) + 1
    ElseIf reverse Then
        endPos = InStrRev(strMain, str2)
        If endPos = 0 Or endPos < startPos Then Exit Function
    Else
        endPos = InStr(startPos, strMain, str2)
        If endPos = 0 Then Exit Function
    End If

    StringBetween = Mid(strMain, startPos, endPos - startPos)
End Function

Public Function SwapString(sInput, sSearch, sReplace As String) As String
    Dim iPos As Integer
    
    iPos = InStr(sInput, sSearch)
    
    If iPos > 0 Then
        SwapString = Left(sInput, iPos - 1) + sReplace + Mid(sInput, iPos + Len(sSearch))
    Else
        SwapString = sInput
    End If
End Function

Public Function Find_Last(sInput, sSearch As String) As Integer
    Find_Last = InStrRev(sInput, sSearch)
End Function

Public Function StringAfterLast(sInput, sSearch As String) As String
    Dim i As Integer
    
    i = Find_Last(sInput, sSearch)
    
    If i = 0 Then
        StringAfterLast = ""
    Else
        StringAfterLast = Mid(sInput, i + Len(sSearch), Len(sInput))
    End If
End Function

Public Function StringBeforeLast(sInput, sSearch As String) As String
    Dim i As Integer
    
    i = Find_Last(sInput, sSearch)
    
    If i = 0 Then
        StringBeforeLast = sInput
    Else
        StringBeforeLast = Mid(sInput, 1, i - 1)
    End If
End Function

' https://stackoverflow.com/questions/10978560/converting-to-sentence-case-using-vba
Public Function SentenceCase(sText As String) As String
    Dim i As Long, bCap As Boolean, ch As String * 1
    
    SentenceCase = LCase(sText)       '-- convert all to lowercase first
    bCap = True
    For i = 1 To Len(SentenceCase)
        ch = Mid$(SentenceCase, i, 1)
        Select Case AscW(ch)
            Case 97 To 122 '-- a-z : separated and put on top as happens more often
                If bCap Then
                    Mid$(SentenceCase, i, 1) = UCase(ch)
                    bCap = False
                End If
            Case 33, 46, 63, 10, 13   '-- sentence terminators ! . ? Lf Cr
                bCap = True
            Case 32, 160, 9           '-- space, non-break space, tab
            Case 34, 41, 93, 125, 148 '-- closing quotes or brackets
            Case Is < 128             '-- other chars between 0-127
                If bCap Then bCap = False
            Case Else                 '-- Extended-Ascii (128-255) or Unicode (> 255)
                If bCap Then
                    If StrComp(ch, UCase(ch), vbBinaryCompare) <> 0 Then
                        '-- a letter that has uppercase.
                        Mid$(SentenceCase, i, 1) = UCase(ch)
                    End If
                    bCap = False
                End If
        End Select
    Next
End Function

Public Sub ConvertSelectedToTitleCase()
    Dim txtOnly As Range
    
    On Error Resume Next
    
    ' only constants that are text
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    
    If Not txtOnly Is Nothing Then
        txtOnly.Value = Evaluate("PROPER(" & txtOnly.Address & ")")
    End If
End Sub

Public Sub ConvertSelectedToSentenceCase()
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim txtOnly As Range
    
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    
    If Not txtOnly Is Nothing Then
        arr = txtOnly.Value
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                arr(r, c) = SentenceCase(CStr(arr(r, c)))
            Next c
        Next r
        txtOnly.Value = arr
    End If
    
    Erase arr
End Sub

Public Sub ConvertSelectedToUpperCase()
    Dim txtOnly As Range
    
    On Error Resume Next
    
    ' only constants that are text
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    
    If Not txtOnly Is Nothing Then
        txtOnly.Value = Evaluate("UPPER(" & txtOnly.Address & ")")
    End If
End Sub

Public Sub ConvertSelectedToLowerCase()
    Dim txtOnly As Range
    
    On Error Resume Next
    
    ' only constants that are text
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    
    If Not txtOnly Is Nothing Then
        txtOnly.Value = Evaluate("LOWER(" & txtOnly.Address & ")")
    End If
End Sub

Function IsLatin(ByVal Str As String) As Boolean
    Dim i As Long
    Dim codePoint As Long
    IsLatin = True
    
    ' RE: µ I know, lazy hack. but it doesn't break the exports...
    For i = 1 To Len(Str)
        codePoint = AscW(Mid(Str, i, 1))
        If Not ((codePoint >= 0 And codePoint <= 255) Or Mid(Str, i, 1) = "µ") Then
            IsLatin = False
            Exit Function
        End If
    Next i
End Function
