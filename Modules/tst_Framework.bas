Attribute VB_Name = "tst_Framework"
' 2025 08 15 - Added unit.  Framework and unit tests developed by copilot/Mike Thompson
'
' Interface :-)
'
' Public ActiveTestModule As String
' Public Sub AssertEqual(TestName As String, Expected As Variant, Actual As Variant)
' Public Sub AssertTrue(TestName As String, Actual As Boolean)
' Public Sub AssertFalse(TestName As String, Actual As Boolean)
' Function CreateTestSheet(sheetName As String) As Worksheet
' Sub DeleteTestSheet(sheetName As String)

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
    Call Test_LibraryString
    Call Test_LibraryMath
    Call Test_LibraryArray
    Call Test_LibraryClipboard
    Call Test_LibraryDate
    Call Test_LibraryFiles
    Call Test_LibraryControls

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

Function CreateTestSheet(sheetName As String) As Worksheet
    Set CreateTestSheet = ThisWorkbook.Sheets.Add
    CreateTestSheet.Name = sheetName
    CreateTestSheet.Activate
End Function

Sub DeleteTestSheet(sheetName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub




