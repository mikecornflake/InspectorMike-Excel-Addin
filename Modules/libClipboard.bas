Attribute VB_Name = "libClipboard"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass
'            - Tests by copilot/ChatGPT-5
'
'            - MSForms deprecated and were causing Excel crashes.  Entirely new Clipboard routines added
'            -      GetClipboard / SetClipboard / Clipboard_Clear
'            -      Stackoverflow / Mike T / copilot joint venture
'            - Defensive coding added to all routines

Option Explicit
Option Private Module

Public Sub Test_LibraryClipboard()
    ActiveTestModule = "LibraryClipboard"

    Dim result As String
    
    ' === Set & Get Clipboard Test ===
    Call Clipboard_Set("21-1-16")
    result = Clipboard_Get
    Call AssertEqual("Get/SetClipboard - 21-1-16", result, "21-1-16")

    ' === Clear Clipboard ===
    Call Clipboard_Clear
    result = Clipboard_Get
    Call AssertEqual("Clipboard_Clear - should result in empty string", vbNullString, result)
End Sub

' https://stackoverflow.com/questions/14219455/excel-vba-code-to-copy-a-specific-string-to-clipboard/60896244#60896244
' The below isn't working in my flavour of VBA
'Function Clipboard$(Optional s$)
'    Dim v: v = s  'Cast to variant for 64-bit VBA support
'    With CreateObject("htmlfile")
'        With .parentWindow.clipboardData
'            Select Case True
'                Case Len(s): .setData "text", v
'                Case Else:   Clipboard = .GetData("text")
'            End Select
'        End With
'    End With
'End Function

' Refactored versions of the above
Function Clipboard_Get() As Variant
    On Error GoTo ErrHandler
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Clipboard_Get = .GetData("text")
        End With
    End With
    Exit Function
ErrHandler:
    MsgBox "Failed to read clipboard: " & Err.Description
    Clipboard_Get = ""
End Function

Function Clipboard_Set(s As Variant)
    On Error GoTo ErrHandler
    Dim v: v = s  ' Cast to variant for 64-bit VBA support
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", v
        End With
    End With
    Exit Function
ErrHandler:
    MsgBox "Failed to set clipboard: " & Err.Description
End Function

' 2025 08 15 - Added by copilot
' 2026 04 24 - Simplified by chatgpt 5.4
Public Sub Clipboard_Clear()
    Clipboard_Set vbNullString
End Sub
