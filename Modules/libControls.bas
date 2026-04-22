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

Public Sub Test_LibraryControls()
    ActiveTestModule = "libControls"
    
    Test_IsSupportedControlType
    Test_ComboContains
    Test_RemoveComboItem
    Test_GetControlValue
    Test_SetControlValue
    Test_RoundTrip
    Test_ControlTypeConsistency
End Sub

' ComboBox Item management
Public Sub RemoveComboItem(ByVal cbo As MSForms.ComboBox, ByVal pValue As String)
    Dim i As Long
    
    For i = cbo.ListCount - 1 To 0 Step -1
        If StrComp(cbo.List(i), pValue, vbTextCompare) = 0 Then
            cbo.RemoveItem i
            Exit For
        End If
    Next i
End Sub

' ComboBox Item management
Public Function ComboContains(ByVal cbo As MSForms.ComboBox, ByVal pValue As String) As Boolean
    Dim i As Long
    
    ComboContains = False
    
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), pValue, vbTextCompare) = 0 Then
            ComboContains = True
            Exit Function
        End If
    Next i
End Function

' List of controls supported by this module
Public Function IsSupportedControlType(ByVal pControlType As String) As Boolean
    Select Case LCase$(Trim$(pControlType))
        Case "textbox", "combobox", "checkbox"
            IsSupportedControlType = True
        Case Else
            IsSupportedControlType = False
    End Select
End Function

Private Sub RaiseUnsupportedControlError(ByVal pCtl As MSForms.control, ByVal pRoutineName As String)
    Err.Raise _
        Number:=ERR_UNSUPPORTED_CONTROL, _
        source:="LibraryForms." & pRoutineName, _
        Description:="Unsupported control type: " & TypeName(pCtl)
End Sub

' Attempt to get the value of a control
Public Function GetControlValue(ByVal pCtl As MSForms.control) As Variant
    Dim sType As String
    
    sType = TypeName(pCtl)
    
    If Not IsSupportedControlType(sType) Then
        RaiseUnsupportedControlError pCtl, "GetControlValue"
    End If
    
    Select Case sType
        Case "TextBox"
            GetControlValue = Trim$(CStr(pCtl.Text))
        
        Case "ComboBox"
            GetControlValue = Trim$(CStr(pCtl.Value))
        
        Case "CheckBox"
            GetControlValue = CBool(pCtl.Value)
    End Select
End Function

' Attempt to set the value of a control
Public Sub SetControlValue(ByVal pCtl As MSForms.control, ByVal pValue As Variant)
    Dim sType As String
    
    sType = TypeName(pCtl)
    
    If Not IsSupportedControlType(sType) Then
        RaiseUnsupportedControlError pCtl, "SetControlValue"
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

Private Sub ResetTestControls()
    With frmTestControls
        .edtTest.Text = vbNullString
        .cboTest.Clear
        .cboTest.Value = vbNullString
        .cbTest.Value = False
    End With
End Sub

' =========================
' IsSupportedControlType
' =========================
Private Sub Test_IsSupportedControlType()
    Call AssertTrue("IsSupportedControlType - TextBox", IsSupportedControlType("TextBox"))
    Call AssertTrue("IsSupportedControlType - textbox", IsSupportedControlType("textbox"))
    Call AssertTrue("IsSupportedControlType - padded textbox", IsSupportedControlType(" TextBox "))
    
    Call AssertTrue("IsSupportedControlType - ComboBox", IsSupportedControlType("ComboBox"))
    Call AssertTrue("IsSupportedControlType - combobox", IsSupportedControlType("combobox"))
    Call AssertTrue("IsSupportedControlType - padded combobox", IsSupportedControlType(" ComboBox "))
    
    Call AssertTrue("IsSupportedControlType - CheckBox", IsSupportedControlType("CheckBox"))
    Call AssertTrue("IsSupportedControlType - checkbox", IsSupportedControlType("checkbox"))
    Call AssertTrue("IsSupportedControlType - padded checkbox", IsSupportedControlType(" CheckBox "))
    
    Call AssertFalse("IsSupportedControlType - Combo", IsSupportedControlType("Combo"))
    Call AssertFalse("IsSupportedControlType - Text", IsSupportedControlType("Text"))
    Call AssertFalse("IsSupportedControlType - Label", IsSupportedControlType("Label"))
    Call AssertFalse("IsSupportedControlType - empty", IsSupportedControlType(vbNullString))
    Call AssertFalse("IsSupportedControlType - spaces", IsSupportedControlType("   "))
End Sub

' =========================
' ComboContains
' =========================
Private Sub Test_ComboContains()
    ResetTestControls
    
    Call AssertFalse("ComboContains - empty combo", ComboContains(frmTestControls.cboTest, "apple"))
    
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    Call AssertTrue("ComboContains - exact match", ComboContains(frmTestControls.cboTest, "banana"))
    Call AssertFalse("ComboContains - not found", ComboContains(frmTestControls.cboTest, "grape"))
    Call AssertTrue("ComboContains - case insensitive upper", ComboContains(frmTestControls.cboTest, "BANANA"))
    Call AssertTrue("ComboContains - case insensitive mixed", ComboContains(frmTestControls.cboTest, "bAnAnA"))
    
    Call AssertFalse("ComboContains - leading space literal", ComboContains(frmTestControls.cboTest, " banana"))
    Call AssertFalse("ComboContains - trailing space literal", ComboContains(frmTestControls.cboTest, "banana "))
    
    frmTestControls.cboTest.AddItem "banana"
    Call AssertTrue("ComboContains - duplicates", ComboContains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    frmTestControls.cboTest.AddItem vbNullString
    Call AssertTrue("ComboContains - blank item", ComboContains(frmTestControls.cboTest, vbNullString))
End Sub

' =========================
' RemoveComboItem
' =========================
Private Sub Test_RemoveComboItem()
    Dim actual As String
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    RemoveComboItem frmTestControls.cboTest, "banana"
    Call AssertEqual("RemoveComboItem - remove existing count", 2, frmTestControls.cboTest.ListCount)
    Call AssertFalse("RemoveComboItem - removed item absent", ComboContains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    RemoveComboItem frmTestControls.cboTest, "grape"
    Call AssertEqual("RemoveComboItem - remove missing count unchanged", 3, frmTestControls.cboTest.ListCount)
    Call AssertEqual("RemoveComboItem - remove missing item0", "apple", frmTestControls.cboTest.List(0))
    Call AssertEqual("RemoveComboItem - remove missing item1", "banana", frmTestControls.cboTest.List(1))
    Call AssertEqual("RemoveComboItem - remove missing item2", "cherry", frmTestControls.cboTest.List(2))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    RemoveComboItem frmTestControls.cboTest, "BANANA"
    Call AssertEqual("RemoveComboItem - case insensitive count", 2, frmTestControls.cboTest.ListCount)
    Call AssertFalse("RemoveComboItem - case insensitive removed", ComboContains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    frmTestControls.cboTest.AddItem "apple"
    RemoveComboItem frmTestControls.cboTest, " apple"
    Call AssertEqual("RemoveComboItem - whitespace literal count", 1, frmTestControls.cboTest.ListCount)
    Call AssertEqual("RemoveComboItem - whitespace literal unchanged", "apple", frmTestControls.cboTest.List(0))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    RemoveComboItem frmTestControls.cboTest, "banana"
    Call AssertEqual("RemoveComboItem - duplicates count reduced by one", 3, frmTestControls.cboTest.ListCount)
    Call AssertTrue("RemoveComboItem - duplicates one remains", ComboContains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "pear"
        .AddItem "banana"
    End With
    
    RemoveComboItem frmTestControls.cboTest, "banana"
    actual = frmTestControls.cboTest.List(0) & "|" & _
             frmTestControls.cboTest.List(1) & "|" & _
             frmTestControls.cboTest.List(2)
    Call AssertEqual("RemoveComboItem - removes last matching occurrence", "apple|banana|pear", actual)
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem vbNullString
        .AddItem "banana"
    End With
    
    RemoveComboItem frmTestControls.cboTest, vbNullString
    Call AssertEqual("RemoveComboItem - remove blank item count", 2, frmTestControls.cboTest.ListCount)
    Call AssertEqual("RemoveComboItem - remove blank item item0", "apple", frmTestControls.cboTest.List(0))
    Call AssertEqual("RemoveComboItem - remove blank item item1", "banana", frmTestControls.cboTest.List(1))
End Sub

' =========================
' GetControlValue
' =========================
Private Sub Test_GetControlValue()
    ResetTestControls
    
    frmTestControls.edtTest.Text = "hello"
    Call AssertEqual("GetControlValue - textbox simple", "hello", GetControlValue(frmTestControls.edtTest))
    
    frmTestControls.edtTest.Text = "  hello  "
    Call AssertEqual("GetControlValue - textbox trimmed", "hello", GetControlValue(frmTestControls.edtTest))
    
    frmTestControls.edtTest.Text = vbNullString
    Call AssertEqual("GetControlValue - textbox blank", vbNullString, GetControlValue(frmTestControls.edtTest))
    
    frmTestControls.edtTest.Text = "   "
    Call AssertEqual("GetControlValue - textbox spaces", vbNullString, GetControlValue(frmTestControls.edtTest))
    
    frmTestControls.cboTest.Value = "hello"
    Call AssertEqual("GetControlValue - combobox simple", "hello", GetControlValue(frmTestControls.cboTest))
    
    frmTestControls.cboTest.Value = "  hello  "
    Call AssertEqual("GetControlValue - combobox trimmed", "hello", GetControlValue(frmTestControls.cboTest))
    
    frmTestControls.cboTest.Value = vbNullString
    Call AssertEqual("GetControlValue - combobox blank", vbNullString, GetControlValue(frmTestControls.cboTest))
    
    frmTestControls.cboTest.Value = "   "
    Call AssertEqual("GetControlValue - combobox spaces", vbNullString, GetControlValue(frmTestControls.cboTest))
    
    frmTestControls.cbTest.Value = True
    Call AssertTrue("GetControlValue - checkbox true", CBool(GetControlValue(frmTestControls.cbTest)))
    
    frmTestControls.cbTest.Value = False
    Call AssertFalse("GetControlValue - checkbox false", CBool(GetControlValue(frmTestControls.cbTest)))
End Sub

' =========================
' SetControlValue
' =========================
Private Sub Test_SetControlValue()
    ResetTestControls
    
    SetControlValue frmTestControls.edtTest, "hello"
    Call AssertEqual("SetControlValue - textbox simple", "hello", frmTestControls.edtTest.Text)
    
    SetControlValue frmTestControls.edtTest, "  hello  "
    Call AssertEqual("SetControlValue - textbox trimmed", "hello", frmTestControls.edtTest.Text)
    
    SetControlValue frmTestControls.edtTest, vbNullString
    Call AssertEqual("SetControlValue - textbox blank", vbNullString, frmTestControls.edtTest.Text)
    
    SetControlValue frmTestControls.edtTest, "   "
    Call AssertEqual("SetControlValue - textbox spaces", vbNullString, frmTestControls.edtTest.Text)
    
    SetControlValue frmTestControls.edtTest, Null
    Call AssertEqual("SetControlValue - textbox null", vbNullString, frmTestControls.edtTest.Text)
    
    SetControlValue frmTestControls.edtTest, 123
    Call AssertEqual("SetControlValue - textbox numeric", "123", frmTestControls.edtTest.Text)
    
    SetControlValue frmTestControls.edtTest, True
    Call AssertEqual("SetControlValue - textbox boolean", "True", frmTestControls.edtTest.Text)
    
    SetControlValue frmTestControls.cboTest, "hello"
    Call AssertEqual("SetControlValue - combobox simple", "hello", frmTestControls.cboTest.Value)
    
    SetControlValue frmTestControls.cboTest, "  hello  "
    Call AssertEqual("SetControlValue - combobox trimmed", "hello", frmTestControls.cboTest.Value)
    
    SetControlValue frmTestControls.cboTest, vbNullString
    Call AssertEqual("SetControlValue - combobox blank", vbNullString, frmTestControls.cboTest.Value)
    
    SetControlValue frmTestControls.cboTest, "   "
    Call AssertEqual("SetControlValue - combobox spaces", vbNullString, frmTestControls.cboTest.Value)
    
    SetControlValue frmTestControls.cboTest, Null
    Call AssertEqual("SetControlValue - combobox null", vbNullString, frmTestControls.cboTest.Value)
    
    SetControlValue frmTestControls.cboTest, 123
    Call AssertEqual("SetControlValue - combobox numeric", "123", frmTestControls.cboTest.Value)
    
    SetControlValue frmTestControls.cboTest, True
    Call AssertEqual("SetControlValue - combobox boolean", "True", frmTestControls.cboTest.Value)
    
    SetControlValue frmTestControls.cbTest, "true"
    Call AssertTrue("SetControlValue - checkbox true text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "TRUE"
    Call AssertTrue("SetControlValue - checkbox TRUE text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, " yes "
    Call AssertTrue("SetControlValue - checkbox yes text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "y"
    Call AssertTrue("SetControlValue - checkbox y text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "1"
    Call AssertTrue("SetControlValue - checkbox 1 text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "false"
    Call AssertFalse("SetControlValue - checkbox false text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "FALSE"
    Call AssertFalse("SetControlValue - checkbox FALSE text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, " no "
    Call AssertFalse("SetControlValue - checkbox no text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "n"
    Call AssertFalse("SetControlValue - checkbox n text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "0"
    Call AssertFalse("SetControlValue - checkbox 0 text", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, Null
    Call AssertFalse("SetControlValue - checkbox null", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, vbNullString
    Call AssertFalse("SetControlValue - checkbox blank", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, "   "
    Call AssertFalse("SetControlValue - checkbox spaces", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, True
    Call AssertTrue("SetControlValue - checkbox boolean true", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, False
    Call AssertFalse("SetControlValue - checkbox boolean false", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, 1
    Call AssertTrue("SetControlValue - checkbox numeric 1", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, 0
    Call AssertFalse("SetControlValue - checkbox numeric 0", frmTestControls.cbTest.Value)
    
    SetControlValue frmTestControls.cbTest, -1
    Call AssertTrue("SetControlValue - checkbox numeric -1", frmTestControls.cbTest.Value)
    
    Test_SetControlValue_CheckBoxInvalidValue "maybe", "SetControlValue - checkbox invalid maybe"
    Test_SetControlValue_CheckBoxInvalidValue "abc", "SetControlValue - checkbox invalid abc"
    
    SetControlValue frmTestControls.cbTest, 2
    Call AssertTrue("SetControlValue - checkbox numeric 2", frmTestControls.cbTest.Value)
End Sub

Private Sub Test_SetControlValue_CheckBoxInvalidValue(ByVal pValue As Variant, ByVal pTestName As String)
    On Error Resume Next
    Err.Clear
    
    frmTestControls.cbTest.Value = False
    SetControlValue frmTestControls.cbTest, pValue
    
    Call AssertTrue(pTestName, Err.Number <> 0)
    
    On Error GoTo 0
End Sub

' =========================
' Round trip tests
' =========================
Private Sub Test_RoundTrip()
    ResetTestControls
    
    SetControlValue frmTestControls.edtTest, "  hello  "
    Call AssertEqual("RoundTrip - textbox", "hello", GetControlValue(frmTestControls.edtTest))
    
    SetControlValue frmTestControls.cboTest, "  hello  "
    Call AssertEqual("RoundTrip - combobox", "hello", GetControlValue(frmTestControls.cboTest))
    
    SetControlValue frmTestControls.cbTest, "yes"
    Call AssertTrue("RoundTrip - checkbox true", CBool(GetControlValue(frmTestControls.cbTest)))
    
    SetControlValue frmTestControls.cbTest, "no"
    Call AssertFalse("RoundTrip - checkbox false", CBool(GetControlValue(frmTestControls.cbTest)))
End Sub

' =========================
' Consistency checks
' =========================
Private Sub Test_ControlTypeConsistency()
    Call AssertTrue("ControlTypeConsistency - TextBox typename supported", _
                    IsSupportedControlType(TypeName(frmTestControls.edtTest)))
    
    Call AssertTrue("ControlTypeConsistency - ComboBox typename supported", _
                    IsSupportedControlType(TypeName(frmTestControls.cboTest)))
    
    Call AssertTrue("ControlTypeConsistency - CheckBox typename supported", _
                    IsSupportedControlType(TypeName(frmTestControls.cbTest)))
End Sub
