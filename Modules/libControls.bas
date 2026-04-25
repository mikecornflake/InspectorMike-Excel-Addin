Attribute VB_Name = "libControls"
Option Explicit
Option Private Module

'
' All routines in here for use with Excel VBA Forms
'
' All routines in here should be kept in sync.
' If you extend one function, you likely need to extend the others
'

Private Const ERR_BASE_LIBRARY_FORMS As Long = vbObjectError + 2048
Private Const ERR_UNSUPPORTED_CONTROL As Long = ERR_BASE_LIBRARY_FORMS + 1

Private Const DATE_FORMAT As String = "yyyy-mm-dd"
Private Const TIME_FORMAT As String = "HH:nn:ss"
Private Const DATETIME_FORMAT As String = "yyyy-mm-dd HH:nn:ss"

Private Sub Control_RaiseUnsupportedError(ByVal pCtl As MSForms.control, ByVal pRoutineName As String)
    Err.Raise _
        Number:=ERR_UNSUPPORTED_CONTROL, _
        source:="LibraryForms." & pRoutineName, _
        Description:="Unsupported control type: " & TypeName(pCtl)
End Sub

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
        Case "textbox", "combobox", "checkbox", "date", "time", "datetime"
            Control_IsSupportedType = True
        Case Else
            Control_IsSupportedType = False
    End Select
End Function

Public Function Control_GetLogicalType(ByVal pCtl As MSForms.control) As String
    Dim sName As String
    
    sName = LCase$(pCtl.Name)
    
    Select Case Left$(sName, 4)
        Case "dat_"
            Control_GetLogicalType = "date"
            Exit Function
        
        Case "tim_"
            Control_GetLogicalType = "time"
            Exit Function
        
        Case "dtm_"
            Control_GetLogicalType = "datetime"
            Exit Function
    End Select
    
    Control_GetLogicalType = LCase$(TypeName(pCtl))
End Function

' Attempt to get the value of a control
Public Function Control_GetValue(ByVal pCtl As MSForms.control) As Variant
    Dim sType As String
    Dim sValue As String
    
    sType = Control_GetLogicalType(pCtl)
    
    If Not Control_IsSupportedType(sType) Then
        Control_RaiseUnsupportedError pCtl, "Control_GetValue"
    End If
    
    Select Case sType
        Case "textbox"
            Control_GetValue = Trim$(CStr(pCtl.Text))
        
        Case "combobox"
            Control_GetValue = Trim$(CStr(pCtl.Value))
        
        Case "checkbox"
            Control_GetValue = CBool(pCtl.Value)
        
        Case "date"
            sValue = Trim$(CStr(pCtl.Text))
            If Len(sValue) = 0 Then
                Control_GetValue = vbNullString
            ElseIf IsDate(sValue) Then
                Control_GetValue = Format$(CDate(sValue), DATE_FORMAT)
            Else
                Err.Raise vbObjectError + 513, "Control_GetValue", _
                    "Invalid date: " & sValue
            End If
        
        Case "time"
            sValue = Trim$(CStr(pCtl.Text))
            If Len(sValue) = 0 Then
                Control_GetValue = vbNullString
            ElseIf IsDate(sValue) Then
                Control_GetValue = Format$(TimeValue(sValue), TIME_FORMAT)
            Else
                Err.Raise vbObjectError + 514, "Control_GetValue", _
                    "Invalid time: " & sValue
            End If
        
        Case "datetime"
            sValue = Trim$(CStr(pCtl.Text))
            If Len(sValue) = 0 Then
                Control_GetValue = vbNullString
            ElseIf IsDate(sValue) Then
                Control_GetValue = Format$(CDate(sValue), DATETIME_FORMAT)
            Else
                Err.Raise vbObjectError + 515, "Control_GetValue", _
                    "Invalid date/time: " & sValue
            End If
    End Select
End Function

' Attempt to set the value of a control
Public Sub Control_SetValue(ByVal pCtl As MSForms.control, ByVal pValue As Variant)
    Dim sType As String
    
    sType = Control_GetLogicalType(pCtl)
    
    If Not Control_IsSupportedType(sType) Then
        Control_RaiseUnsupportedError pCtl, "Control_SetValue"
    End If
    
    Select Case sType
        Case "textbox"
            If IsNull(pValue) Then
                pCtl.Text = vbNullString
            Else
                pCtl.Text = Trim$(CStr(pValue))
            End If
        
        Case "combobox"
            If IsNull(pValue) Then
                pCtl.Value = vbNullString
            Else
                pCtl.Value = Trim$(CStr(pValue))
            End If
        
        Case "checkbox"
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
        
        Case "date"
            If IsNull(pValue) Or Len(Trim$(CStr(pValue))) = 0 Then
                pCtl.Text = vbNullString
            
            ElseIf IsNumeric(pValue) Then
                pCtl.Text = Format$(CDbl(pValue), DATE_FORMAT)
            
            ElseIf IsDate(pValue) Then
                pCtl.Text = Format$(CDate(pValue), DATE_FORMAT)
            
            Else
                pCtl.Text = Trim$(CStr(pValue))
            End If
        
        Case "time"
            If IsNull(pValue) Or Len(Trim$(CStr(pValue))) = 0 Then
                pCtl.Text = vbNullString
            
            ElseIf IsNumeric(pValue) Then
                pCtl.Text = Format$(CDbl(pValue), TIME_FORMAT)
            
            ElseIf IsDate(pValue) Then
                pCtl.Text = Format$(TimeValue(pValue), TIME_FORMAT)
            
            Else
                pCtl.Text = Trim$(CStr(pValue))
            End If
        
        Case "datetime"
            If IsNull(pValue) Or Len(Trim$(CStr(pValue))) = 0 Then
                pCtl.Text = vbNullString
            
            ElseIf IsNumeric(pValue) Then
                pCtl.Text = Format$(CDbl(pValue), DATETIME_FORMAT)
            
            ElseIf IsDate(pValue) Then
                pCtl.Text = Format$(CDate(pValue), DATETIME_FORMAT)
            
            Else
                pCtl.Text = Trim$(CStr(pValue))
            End If
    End Select
End Sub

Public Function Control_IsBlank(ByVal pCtl As MSForms.control) As Boolean
    Select Case TypeName(pCtl)
        Case "TextBox"
            Control_IsBlank = (Len(Trim$(CStr(pCtl.Text))) = 0)
        Case "ComboBox"
            Control_IsBlank = (Len(Trim$(CStr(pCtl.Value))) = 0)
        Case "CheckBox"
            Control_IsBlank = False
        Case Else
            Control_IsBlank = True
    End Select
End Function

