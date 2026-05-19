Attribute VB_Name = "tst_Framework"
' 2025 08 15 - Added unit.  Framework and unit tests developed by copilot/Mike Thompson
'
' Interface :-)
'
' Public ActiveTestModule As String
' Public Sub AssertEqual(TestName As String, Expected As Variant, Actual As Variant)
' Public Sub AssertTrue(TestName As String, Actual As Boolean)
' Public Sub AssertFalse(TestName As String, Actual As Boolean)
' Public Sub AssertRaises(TestName As String, ExpectedErrNumber As Long, ActualErrNumber As Long)
' Public Function CreateTestSheet(sheetName As String) As Worksheet
' Public Sub DeleteTestSheet(sheetName As String)

Option Explicit
Option Private Module

Private Type TestResult
    Module As String
    Name As String
    Passed As Boolean
    Message As String
End Type

Private TestResults() As TestResult
Private TestCount As Long

Public ActiveTestModule As String

Public Sub RunAllTests()
    TestCount = 0
    Erase TestResults
    ActiveTestModule = ""

    ' Determine which tests to run
    ' Library tests
    Call Test_libString
    Call Test_libMath
    Call Test_libArray
    Call Test_libDate
    Call Test_libInterpolation
    
    ' File System Tests
    Call Test_libClipboard
    Call Test_libFiles
    
    ' UI Tests
    Call Test_libControls
    Call Test_libTable

    ' Report results
    Dim i As Long
    Dim bFail As Boolean
    
    bFail = False
    
    Debug.Print ""
    Debug.Print ""
    
    Debug.Print "----- Test Results -----"
    For i = 1 To TestCount
        With TestResults(i)
            bFail = bFail Or Not .Passed
            
            Debug.Print IIf(.Passed, "PASS", "FAIL") & ": Function " & .Module & "." & .Name & ": " & .Message
        End With
    Next i
    Debug.Print "------------------------"
    
    If bFail Then
        Debug.Print ""
        Debug.Print "----- Failed Tests -----"
        For i = 1 To TestCount
            With TestResults(i)
                If Not .Passed Then
                    Debug.Print IIf(.Passed, "PASS", "FAIL") & ": Function " & .Name & ": " & .Message
                End If
            End With
        Next i
        Debug.Print "------------------------"
    Else
        Debug.Print "All Tests Passed!"
    End If
End Sub

Public Sub AssertEqual(TestName As String, Expected As Variant, actual As Variant)
    TestCount = TestCount + 1
    ReDim Preserve TestResults(1 To TestCount)

    With TestResults(TestCount)
        .Module = ActiveTestModule
        .Name = TestName
        If Expected = actual Then
            .Passed = True
            .Message = "Expected and received [" & FormatVariant(Expected) & "]"
        Else
            .Passed = False
            .Message = "Expected [" & FormatVariant(Expected) & "], received [" & FormatVariant(actual) & "]"
        End If
    End With
End Sub

Public Sub AssertTrue(TestName As String, actual As Boolean)
    Call AssertEqual(TestName, True, actual)
End Sub

Public Sub AssertFalse(TestName As String, actual As Boolean)
    Call AssertEqual(TestName, False, actual)
End Sub

Public Sub AssertRaises( _
    ByVal TestName As String, _
    ByVal ExpectedErrNumber As Long, _
    ByVal ActualErrNumber As Long)

    Call AssertEqual(TestName, ExpectedErrNumber, ActualErrNumber)
End Sub

Public Function CreateTestSheet(sheetName As String) As Worksheet
    Dim oSheet As Worksheet
    Dim sName As String
    
    sName = Left$("test_" & sheetName, 31)
    
    Set oSheet = Worksheet_Find(ThisWorkbook, sName)
    
    If oSheet Is Nothing Then
        Set CreateTestSheet = ThisWorkbook.Sheets.Add
        CreateTestSheet.Name = sName
        CreateTestSheet.Activate
    Else
        Set CreateTestSheet = oSheet
    End If
End Function

Public Sub DeleteTestSheet(sheetName As String)
    Dim sName As String
    sName = Left$("test_" & sheetName, 31)
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Function FormatVariant(v As Variant) As String
    If IsError(v) Then
        FormatVariant = "Error #" & CStr(v)
    ElseIf IsNull(v) Then
        FormatVariant = "Null"
    ElseIf isEmpty(v) Then
        FormatVariant = "Empty"
    Else
        FormatVariant = CStr(v)
    End If
End Function





