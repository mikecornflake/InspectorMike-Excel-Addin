Attribute VB_Name = "libString"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass

Option Explicit

Private Const ERR_BASE_LIBRARY_STRING As Long = vbObjectError + 1024
Private Const ERR_INVALID_STRING_BOOL As Long = ERR_BASE_LIBRARY_STRING + 1

Public Function Text_ToBool(ByVal AInput As String) As Boolean
    Dim sInput As String
    
    sInput = LCase$(Trim$(AInput))
    
    Select Case sInput
        Case "true", "yes", "y", "1"
            Text_ToBool = True
        
        Case "false", "no", "n", "0"
            Text_ToBool = False
        
        Case Else
            Err.Raise _
                Number:=ERR_INVALID_STRING_BOOL, _
                source:="LibraryString.Text_ToBool", _
                Description:="Invalid boolean string: [" & AInput & "]"
    End Select
End Function

Public Function Bool_ToText(ByVal AInput As Boolean) As String
    If AInput Then
        Bool_ToText = "True"
    Else
        Bool_ToText = "False"
    End If
End Function

Public Function Text_IsBool(ByVal AInput As String) As Boolean
    Dim sInput As String
    
    sInput = LCase$(Trim$(AInput))
    
    Select Case sInput
        Case "true", "yes", "y", "1", _
             "false", "no", "n", "0"
            Text_IsBool = True
        
        Case Else
            Text_IsBool = False
    End Select
End Function

Public Function Text_Remove(ByVal sInput As String, ByVal sSubString As String) As String
    Dim iSub As Long
    Dim iLen As Long
    
    iSub = InStr(sInput, sSubString)
    iLen = Len(sSubString)
    
    If iSub <> 0 Then
        Text_Remove = Left(sInput, iSub - 1) & Mid(sInput, iSub + iLen)
    Else
        Text_Remove = sInput
    End If
End Function

Public Function Text_IsNumber(ByVal sValue As String) As Boolean
    On Error GoTo ErrorHandler
    
    sValue = Trim(sValue)
    
    Text_IsNumber = ("" & Val(sValue)) = sValue
    
    Exit Function
ErrorHandler:
    On Error Resume Next
    Text_IsNumber = False
End Function

Function Text_IsLatin(ByVal sText As String) As Boolean
    Dim i As Long
    Dim codePoint As Long
    Text_IsLatin = True
    
    ' RE: µ I know, lazy hack. but it doesn't break the exports...
    For i = 1 To Len(sText)
        codePoint = AscW(Mid(sText, i, 1))
        If Not ((codePoint >= 0 And codePoint <= 255) Or Mid(sText, i, 1) = "µ") Then
            Text_IsLatin = False
            Exit Function
        End If
    Next i
End Function

' Mikes version replaced by StackOverflow
' StackOverflows version replaced by copilots :-)
' Developer: copilot.  15/08/2025
Public Function Text_Between(ByVal sText As String, ByVal sStart As String, ByVal sEnd As String, Optional ByVal pReverse As Boolean = False) As String
    Dim startPos As Long, endPos As Long

    ' Handle start delimiter
    If sStart = "" Then
        startPos = 1
    ElseIf pReverse Then
        startPos = InStrRev(sText, sStart)
        If startPos = 0 Then Exit Function
        startPos = startPos + Len(sStart)
    Else
        startPos = InStr(sText, sStart)
        If startPos = 0 Then Exit Function
        startPos = startPos + Len(sStart)
    End If

    ' Handle end delimiter
    If sEnd = "" Then
        endPos = Len(sText) + 1
    ElseIf pReverse Then
        endPos = InStrRev(sText, sEnd)
        If endPos = 0 Or endPos < startPos Then Exit Function
    Else
        endPos = InStr(startPos, sText, sEnd)
        If endPos = 0 Then Exit Function
    End If

    Text_Between = Mid(sText, startPos, endPos - startPos)
End Function

Public Function Text_Replace(ByVal sText As String, ByVal sFind As String, ByVal sReplace As String) As String
    Dim iPos As Long
    
    iPos = InStr(sText, sFind)
    
    If iPos > 0 Then
        Text_Replace = Left(sText, iPos - 1) + sReplace + Mid(sText, iPos + Len(sFind))
    Else
        Text_Replace = sText
    End If
End Function

Public Function Text_FindLast(ByVal sText As String, ByVal sFind As String) As Long
    Text_FindLast = InStrRev(sText, sFind)
End Function

Public Function Text_AfterLast(ByVal sText As String, ByVal sFind As String) As String
    Dim i As Long
    
    i = Text_FindLast(sText, sFind)
    
    If i = 0 Then
        Text_AfterLast = ""
    Else
        Text_AfterLast = Mid(sText, i + Len(sFind), Len(sText))
    End If
End Function

Public Function Text_BeforeLast(ByVal sText As String, ByVal sFind As String) As String
    Dim i As Long
    
    i = Text_FindLast(sText, sFind)
    
    If i = 0 Then
        Text_BeforeLast = sText
    Else
        Text_BeforeLast = Mid(sText, 1, i - 1)
    End If
End Function

' https://stackoverflow.com/questions/10978560/converting-to-sentence-case-using-vba
Public Function Text_ToSentenceCase(ByVal sText As String) As String
    Dim i As Long, bCap As Boolean, ch As String * 1
    
    Text_ToSentenceCase = LCase(sText)       '-- convert all to lowercase first
    bCap = True
    For i = 1 To Len(Text_ToSentenceCase)
        ch = Mid$(Text_ToSentenceCase, i, 1)
        Select Case AscW(ch)
            Case 97 To 122 '-- a-z : separated and put on top as happens more often
                If bCap Then
                    Mid$(Text_ToSentenceCase, i, 1) = UCase(ch)
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
                        Mid$(Text_ToSentenceCase, i, 1) = UCase(ch)
                    End If
                    bCap = False
                End If
        End Select
    Next
End Function

Public Sub Text_TitleCase_Selection()
    Dim txtOnly As Range
    
    On Error Resume Next
    
    ' only constants that are text
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    
    If Not txtOnly Is Nothing Then
        txtOnly.Value = Evaluate("PROPER(" & txtOnly.Address & ")")
    End If
End Sub

Public Sub Text_SentenceCase_Selection()
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim txtOnly As Range
    
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    
    If Not txtOnly Is Nothing Then
        arr = txtOnly.Value
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                arr(r, c) = Text_ToSentenceCase(CStr(arr(r, c)))
            Next c
        Next r
        txtOnly.Value = arr
    End If
    
    Erase arr
End Sub

Public Sub Text_Upper_Selection()
    Dim txtOnly As Range
    
    On Error Resume Next
    
    ' only constants that are text
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    
    If Not txtOnly Is Nothing Then
        txtOnly.Value = Evaluate("UPPER(" & txtOnly.Address & ")")
    End If
End Sub

Public Sub Text_Lower_Selection()
    Dim txtOnly As Range
    
    On Error Resume Next
    
    ' only constants that are text
    Set txtOnly = Selection.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    
    If Not txtOnly Is Nothing Then
        txtOnly.Value = Evaluate("LOWER(" & txtOnly.Address & ")")
    End If
End Sub
