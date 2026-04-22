Attribute VB_Name = "libClipboard"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass
'            - Tests by copilot/ChatGPT-5
'
'            - MSForms deprecated and were causing Excel crashes.  Entirely new Clipboard routines added
'            -      GetClipboard / SetClipboard / Clipboard_Clear
'            -      Stackoverflow / Mike T / copilot joint venture
'            - Defensive coding added to all routines

Option Explicit

Public Sub Test_LibraryClipboard()
    ActiveTestModule = "LibraryClipboard"

    ' === Setup: Simulate clipboard input ===
    Call SetClipboard("21-1-16")

    ' === Run conversion ===
    Call DoConvertDateOnClipboard

    ' === Validate result ===
    Dim result As String
    result = GetClipboard
    
    Call AssertEqual("DoConvertDateOnClipboard - basic conversion", "2016-01-21", result)

    ' === Clear Clipboard ===
    Call Clipboard_Clear

    ' === Validate ===
    result = GetClipboard

    Call AssertEqual("Clipboard_Clear - should result in empty string", "", result)
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
Function GetClipboard() As Variant
    On Error GoTo ErrHandler
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            GetClipboard = .GetData("text")
        End With
    End With
    Exit Function
ErrHandler:
    MsgBox "Failed to read clipboard: " & Err.Description
    GetClipboard = ""
End Function

Function SetClipboard(s As Variant)
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
Sub Clipboard_Clear()
    On Error GoTo ErrHandler
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            .setData "text", ""
        End With
    End With
    Exit Sub
ErrHandler:
    MsgBox "Failed to clear clipboard: " & Err.Description
End Sub

' 2025 08 15 - Copilot added defensive code.  Mike T converted to new Clipboard routines
Public Sub DoConvertDateOnClipboard()
    ' Input dd-mm-yy or dd-mm-yyyy:
    '   21-1-16
    '
    ' Output: yyyy-mm-dd
    
    Dim sInput As String
    Dim sY As String, sM As String, sd As String
    Dim iY As Long, iM As Long, iD As Long
    Dim i1 As Long, i2 As Long
    
    On Error GoTo ErrHandler
    
    sInput = GetClipboard
    
    i1 = InStr(sInput, "-")
    i2 = InStr(i1 + 1, sInput, "-")
    
    If i1 = 0 Or i2 = 0 Then
        MsgBox "Invalid date format"
        Exit Sub
    End If
    
    sd = Trim(Left(sInput, i1 - 1))
    sM = Trim(Mid(sInput, i1 + 1, i2 - i1 - 1))
    sY = Trim(Mid(sInput, i2 + 1, 99))
    
    iY = Val(sY)
    iM = Val(sM)
    iD = Val(sd)
    
    If iY < 2000 Then
        iY = iY + 2000
    End If
    
    If iM > 12 Then
        MsgBox "Error"
        Exit Sub
    End If
    
    sY = Format(iY, "0000")
    sM = Format(iM, "00")
    sd = Format(iD, "00")
    
    Call SetClipboard(sY & "-" & sM & "-" & sd)
    
    Exit Sub
ErrHandler:
    MsgBox "Unexpected error: " & Err.Description
End Sub
