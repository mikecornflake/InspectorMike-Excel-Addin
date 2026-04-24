Attribute VB_Name = "libControls"
Option Explicit

'
' All routines in here for use with Excel VBA Forms
'
' All routines in here should be kept in sync.
' If you extend one function, you likely need to extend the others
'

Private Const ERR_BASE_LIBRARY_FORMS As Long = vbObjectError + 2048
Private Const ERR_UNSUPPORTED_CONTROL As Long = ERR_BASE_LIBRARY_FORMS + 1

' ComboBox Item management
Public Sub ComboBox_RemoveItem(ByVal cbo As MSForms.ComboBox, ByVal pValue As String)
    Dim i As Long
    
    For i = cbo.ListCount - 1 To 0 Step -1
        If StrComp(cbo.List(i), pValue, vbTextCompare) = 0 Then
            cbo.RemoveItem i
            Exit For
        End If
    Next i
End Sub

' ComboBox Item management
Public Function ComboBox_Contains(ByVal cbo As MSForms.ComboBox, ByVal pValue As String) As Boolean
    Dim i As Long
    
    ComboBox_Contains = False
    
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), pValue, vbTextCompare) = 0 Then
            ComboBox_Contains = True
            Exit Function
        End If
    Next i
End Function

' List of controls supported by this module
Public Function Control_IsSupportedType(ByVal pControlType As String) As Boolean
    Select Case LCase$(Trim$(pControlType))
        Case "textbox", "combobox", "checkbox"
            Control_IsSupportedType = True
        Case Else
            Control_IsSupportedType = False
    End Select
End Function

Private Sub Control_RaiseUnsupportedError(ByVal pCtl As MSForms.control, ByVal pRoutineName As String)
    Err.Raise _
        Number:=ERR_UNSUPPORTED_CONTROL, _
        source:="LibraryForms." & pRoutineName, _
        Description:="Unsupported control type: " & TypeName(pCtl)
End Sub

' Attempt to get the value of a control
Public Function Control_GetValue(ByVal pCtl As MSForms.control) As Variant
    Dim sType As String
    
    sType = TypeName(pCtl)
    
    If Not Control_IsSupportedType(sType) Then
        Control_RaiseUnsupportedError pCtl, "Control_GetValue"
    End If
    
    Select Case sType
        Case "TextBox"
            Control_GetValue = Trim$(CStr(pCtl.Text))
        
        Case "ComboBox"
            Control_GetValue = Trim$(CStr(pCtl.Value))
        
        Case "CheckBox"
            Control_GetValue = CBool(pCtl.Value)
    End Select
End Function

' Attempt to set the value of a control
Public Sub Control_SetValue(ByVal pCtl As MSForms.control, ByVal pValue As Variant)
    Dim sType As String
    
    sType = TypeName(pCtl)
    
    If Not Control_IsSupportedType(sType) Then
        Control_RaiseUnsupportedError pCtl, "Control_SetValue"
    End If
    
    Select Case sType
        Case "TextBox"
            If IsNull(pValue) Then
                pCtl.Text = vbNullString
            Else
                pCtl.Text = Trim$(CStr(pValue))
            End If
        
        Case "ComboBox"
            If IsNull(pValue) Then
                pCtl.Value = vbNullString
            Else
                pCtl.Value = Trim$(CStr(pValue))
            End If
        
        Case "CheckBox"
            If IsNull(pValue) Then
                pCtl.Value = False
            ElseIf Len(Trim$(CStr(pValue))) = 0 Then
                pCtl.Value = False
            Else
                Select Case LCase$(Trim$(CStr(pValue)))
                    Case "true", "yes", "y", "1"
                        pCtl.Value = True
                    Case "false", "no", "n", "0"
                        pCtl.Value = False
                    Case Else
                        pCtl.Value = CBool(pValue)
                End Select
            End If
    End Select
End Sub

