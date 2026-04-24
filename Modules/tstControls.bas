Attribute VB_Name = "tstControls"
Option Private Module

Public Sub Test_LibraryControls()
    ActiveTestModule = "libControls"
    
    Test_Control_IsSupportedType
    Test_ComboBox_Contains
    Test_RemoveComboItem
    Test_Control_GetValue
    Test_Control_SetValue
    Test_RoundTrip
    Test_ControlTypeConsistency
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
' Control_IsSupportedType
' =========================
Private Sub Test_Control_IsSupportedType()
    Call AssertTrue("Control_IsSupportedType - TextBox", Control_IsSupportedType("TextBox"))
    Call AssertTrue("Control_IsSupportedType - textbox", Control_IsSupportedType("textbox"))
    Call AssertTrue("Control_IsSupportedType - padded textbox", Control_IsSupportedType(" TextBox "))
    
    Call AssertTrue("Control_IsSupportedType - ComboBox", Control_IsSupportedType("ComboBox"))
    Call AssertTrue("Control_IsSupportedType - combobox", Control_IsSupportedType("combobox"))
    Call AssertTrue("Control_IsSupportedType - padded combobox", Control_IsSupportedType(" ComboBox "))
    
    Call AssertTrue("Control_IsSupportedType - CheckBox", Control_IsSupportedType("CheckBox"))
    Call AssertTrue("Control_IsSupportedType - checkbox", Control_IsSupportedType("checkbox"))
    Call AssertTrue("Control_IsSupportedType - padded checkbox", Control_IsSupportedType(" CheckBox "))
    
    Call AssertFalse("Control_IsSupportedType - Combo", Control_IsSupportedType("Combo"))
    Call AssertFalse("Control_IsSupportedType - Text", Control_IsSupportedType("Text"))
    Call AssertFalse("Control_IsSupportedType - Label", Control_IsSupportedType("Label"))
    Call AssertFalse("Control_IsSupportedType - empty", Control_IsSupportedType(vbNullString))
    Call AssertFalse("Control_IsSupportedType - spaces", Control_IsSupportedType("   "))
End Sub

' =========================
' ComboBox_Contains
' =========================
Private Sub Test_ComboBox_Contains()
    ResetTestControls
    
    Call AssertFalse("ComboBox_Contains - empty combo", ComboBox_Contains(frmTestControls.cboTest, "apple"))
    
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    Call AssertTrue("ComboBox_Contains - exact match", ComboBox_Contains(frmTestControls.cboTest, "banana"))
    Call AssertFalse("ComboBox_Contains - not found", ComboBox_Contains(frmTestControls.cboTest, "grape"))
    Call AssertTrue("ComboBox_Contains - case insensitive upper", ComboBox_Contains(frmTestControls.cboTest, "BANANA"))
    Call AssertTrue("ComboBox_Contains - case insensitive mixed", ComboBox_Contains(frmTestControls.cboTest, "bAnAnA"))
    
    Call AssertFalse("ComboBox_Contains - leading space literal", ComboBox_Contains(frmTestControls.cboTest, " banana"))
    Call AssertFalse("ComboBox_Contains - trailing space literal", ComboBox_Contains(frmTestControls.cboTest, "banana "))
    
    frmTestControls.cboTest.AddItem "banana"
    Call AssertTrue("ComboBox_Contains - duplicates", ComboBox_Contains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    frmTestControls.cboTest.AddItem vbNullString
    Call AssertTrue("ComboBox_Contains - blank item", ComboBox_Contains(frmTestControls.cboTest, vbNullString))
End Sub

' =========================
' ComboBox_RemoveItem
' =========================
Private Sub Test_RemoveComboItem()
    Dim actual As String
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    ComboBox_RemoveItem frmTestControls.cboTest, "banana"
    Call AssertEqual("ComboBox_RemoveItem - remove existing count", 2, frmTestControls.cboTest.ListCount)
    Call AssertFalse("ComboBox_RemoveItem - removed item absent", ComboBox_Contains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    ComboBox_RemoveItem frmTestControls.cboTest, "grape"
    Call AssertEqual("ComboBox_RemoveItem - remove missing count unchanged", 3, frmTestControls.cboTest.ListCount)
    Call AssertEqual("ComboBox_RemoveItem - remove missing item0", "apple", frmTestControls.cboTest.List(0))
    Call AssertEqual("ComboBox_RemoveItem - remove missing item1", "banana", frmTestControls.cboTest.List(1))
    Call AssertEqual("ComboBox_RemoveItem - remove missing item2", "cherry", frmTestControls.cboTest.List(2))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    ComboBox_RemoveItem frmTestControls.cboTest, "BANANA"
    Call AssertEqual("ComboBox_RemoveItem - case insensitive count", 2, frmTestControls.cboTest.ListCount)
    Call AssertFalse("ComboBox_RemoveItem - case insensitive removed", ComboBox_Contains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    frmTestControls.cboTest.AddItem "apple"
    ComboBox_RemoveItem frmTestControls.cboTest, " apple"
    Call AssertEqual("ComboBox_RemoveItem - whitespace literal count", 1, frmTestControls.cboTest.ListCount)
    Call AssertEqual("ComboBox_RemoveItem - whitespace literal unchanged", "apple", frmTestControls.cboTest.List(0))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "banana"
        .AddItem "cherry"
    End With
    
    ComboBox_RemoveItem frmTestControls.cboTest, "banana"
    Call AssertEqual("ComboBox_RemoveItem - duplicates count reduced by one", 3, frmTestControls.cboTest.ListCount)
    Call AssertTrue("ComboBox_RemoveItem - duplicates one remains", ComboBox_Contains(frmTestControls.cboTest, "banana"))
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem "banana"
        .AddItem "pear"
        .AddItem "banana"
    End With
    
    ComboBox_RemoveItem frmTestControls.cboTest, "banana"
    actual = frmTestControls.cboTest.List(0) & "|" & _
             frmTestControls.cboTest.List(1) & "|" & _
             frmTestControls.cboTest.List(2)
    Call AssertEqual("ComboBox_RemoveItem - removes last matching occurrence", "apple|banana|pear", actual)
    
    ResetTestControls
    With frmTestControls.cboTest
        .AddItem "apple"
        .AddItem vbNullString
        .AddItem "banana"
    End With
    
    ComboBox_RemoveItem frmTestControls.cboTest, vbNullString
    Call AssertEqual("ComboBox_RemoveItem - remove blank item count", 2, frmTestControls.cboTest.ListCount)
    Call AssertEqual("ComboBox_RemoveItem - remove blank item item0", "apple", frmTestControls.cboTest.List(0))
    Call AssertEqual("ComboBox_RemoveItem - remove blank item item1", "banana", frmTestControls.cboTest.List(1))
End Sub

' =========================
' Control_GetValue
' =========================
Private Sub Test_Control_GetValue()
    ResetTestControls
    
    frmTestControls.edtTest.Text = "hello"
    Call AssertEqual("Control_GetValue - textbox simple", "hello", Control_GetValue(frmTestControls.edtTest))
    
    frmTestControls.edtTest.Text = "  hello  "
    Call AssertEqual("Control_GetValue - textbox trimmed", "hello", Control_GetValue(frmTestControls.edtTest))
    
    frmTestControls.edtTest.Text = vbNullString
    Call AssertEqual("Control_GetValue - textbox blank", vbNullString, Control_GetValue(frmTestControls.edtTest))
    
    frmTestControls.edtTest.Text = "   "
    Call AssertEqual("Control_GetValue - textbox spaces", vbNullString, Control_GetValue(frmTestControls.edtTest))
    
    frmTestControls.cboTest.Value = "hello"
    Call AssertEqual("Control_GetValue - combobox simple", "hello", Control_GetValue(frmTestControls.cboTest))
    
    frmTestControls.cboTest.Value = "  hello  "
    Call AssertEqual("Control_GetValue - combobox trimmed", "hello", Control_GetValue(frmTestControls.cboTest))
    
    frmTestControls.cboTest.Value = vbNullString
    Call AssertEqual("Control_GetValue - combobox blank", vbNullString, Control_GetValue(frmTestControls.cboTest))
    
    frmTestControls.cboTest.Value = "   "
    Call AssertEqual("Control_GetValue - combobox spaces", vbNullString, Control_GetValue(frmTestControls.cboTest))
    
    frmTestControls.cbTest.Value = True
    Call AssertTrue("Control_GetValue - checkbox true", CBool(Control_GetValue(frmTestControls.cbTest)))
    
    frmTestControls.cbTest.Value = False
    Call AssertFalse("Control_GetValue - checkbox false", CBool(Control_GetValue(frmTestControls.cbTest)))
End Sub

' =========================
' Control_SetValue
' =========================
Private Sub Test_Control_SetValue()
    ResetTestControls
    
    Control_SetValue frmTestControls.edtTest, "hello"
    Call AssertEqual("Control_SetValue - textbox simple", "hello", frmTestControls.edtTest.Text)
    
    Control_SetValue frmTestControls.edtTest, "  hello  "
    Call AssertEqual("Control_SetValue - textbox trimmed", "hello", frmTestControls.edtTest.Text)
    
    Control_SetValue frmTestControls.edtTest, vbNullString
    Call AssertEqual("Control_SetValue - textbox blank", vbNullString, frmTestControls.edtTest.Text)
    
    Control_SetValue frmTestControls.edtTest, "   "
    Call AssertEqual("Control_SetValue - textbox spaces", vbNullString, frmTestControls.edtTest.Text)
    
    Control_SetValue frmTestControls.edtTest, Null
    Call AssertEqual("Control_SetValue - textbox null", vbNullString, frmTestControls.edtTest.Text)
    
    Control_SetValue frmTestControls.edtTest, 123
    Call AssertEqual("Control_SetValue - textbox numeric", "123", frmTestControls.edtTest.Text)
    
    Control_SetValue frmTestControls.edtTest, True
    Call AssertEqual("Control_SetValue - textbox boolean", "True", frmTestControls.edtTest.Text)
    
    Control_SetValue frmTestControls.cboTest, "hello"
    Call AssertEqual("Control_SetValue - combobox simple", "hello", frmTestControls.cboTest.Value)
    
    Control_SetValue frmTestControls.cboTest, "  hello  "
    Call AssertEqual("Control_SetValue - combobox trimmed", "hello", frmTestControls.cboTest.Value)
    
    Control_SetValue frmTestControls.cboTest, vbNullString
    Call AssertEqual("Control_SetValue - combobox blank", vbNullString, frmTestControls.cboTest.Value)
    
    Control_SetValue frmTestControls.cboTest, "   "
    Call AssertEqual("Control_SetValue - combobox spaces", vbNullString, frmTestControls.cboTest.Value)
    
    Control_SetValue frmTestControls.cboTest, Null
    Call AssertEqual("Control_SetValue - combobox null", vbNullString, frmTestControls.cboTest.Value)
    
    Control_SetValue frmTestControls.cboTest, 123
    Call AssertEqual("Control_SetValue - combobox numeric", "123", frmTestControls.cboTest.Value)
    
    Control_SetValue frmTestControls.cboTest, True
    Call AssertEqual("Control_SetValue - combobox boolean", "True", frmTestControls.cboTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "true"
    Call AssertTrue("Control_SetValue - checkbox true text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "TRUE"
    Call AssertTrue("Control_SetValue - checkbox TRUE text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, " yes "
    Call AssertTrue("Control_SetValue - checkbox yes text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "y"
    Call AssertTrue("Control_SetValue - checkbox y text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "1"
    Call AssertTrue("Control_SetValue - checkbox 1 text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "false"
    Call AssertFalse("Control_SetValue - checkbox false text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "FALSE"
    Call AssertFalse("Control_SetValue - checkbox FALSE text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, " no "
    Call AssertFalse("Control_SetValue - checkbox no text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "n"
    Call AssertFalse("Control_SetValue - checkbox n text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "0"
    Call AssertFalse("Control_SetValue - checkbox 0 text", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, Null
    Call AssertFalse("Control_SetValue - checkbox null", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, vbNullString
    Call AssertFalse("Control_SetValue - checkbox blank", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, "   "
    Call AssertFalse("Control_SetValue - checkbox spaces", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, True
    Call AssertTrue("Control_SetValue - checkbox boolean true", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, False
    Call AssertFalse("Control_SetValue - checkbox boolean false", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, 1
    Call AssertTrue("Control_SetValue - checkbox numeric 1", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, 0
    Call AssertFalse("Control_SetValue - checkbox numeric 0", frmTestControls.cbTest.Value)
    
    Control_SetValue frmTestControls.cbTest, -1
    Call AssertTrue("Control_SetValue - checkbox numeric -1", frmTestControls.cbTest.Value)
    
    Test_Control_SetValue_CheckBoxInvalidValue "maybe", "Control_SetValue - checkbox invalid maybe"
    Test_Control_SetValue_CheckBoxInvalidValue "abc", "Control_SetValue - checkbox invalid abc"
    
    Control_SetValue frmTestControls.cbTest, 2
    Call AssertTrue("Control_SetValue - checkbox numeric 2", frmTestControls.cbTest.Value)
End Sub

Private Sub Test_Control_SetValue_CheckBoxInvalidValue(ByVal pValue As Variant, ByVal pTestName As String)
    On Error Resume Next
    Err.Clear
    
    frmTestControls.cbTest.Value = False
    Control_SetValue frmTestControls.cbTest, pValue
    
    Call AssertTrue(pTestName, Err.Number <> 0)
    
    On Error GoTo 0
End Sub

' =========================
' Round trip tests
' =========================
Private Sub Test_RoundTrip()
    ResetTestControls
    
    Control_SetValue frmTestControls.edtTest, "  hello  "
    Call AssertEqual("RoundTrip - textbox", "hello", Control_GetValue(frmTestControls.edtTest))
    
    Control_SetValue frmTestControls.cboTest, "  hello  "
    Call AssertEqual("RoundTrip - combobox", "hello", Control_GetValue(frmTestControls.cboTest))
    
    Control_SetValue frmTestControls.cbTest, "yes"
    Call AssertTrue("RoundTrip - checkbox true", CBool(Control_GetValue(frmTestControls.cbTest)))
    
    Control_SetValue frmTestControls.cbTest, "no"
    Call AssertFalse("RoundTrip - checkbox false", CBool(Control_GetValue(frmTestControls.cbTest)))
End Sub

' =========================
' Consistency checks
' =========================
Private Sub Test_ControlTypeConsistency()
    Call AssertTrue("ControlTypeConsistency - TextBox typename supported", _
                    Control_IsSupportedType(TypeName(frmTestControls.edtTest)))
    
    Call AssertTrue("ControlTypeConsistency - ComboBox typename supported", _
                    Control_IsSupportedType(TypeName(frmTestControls.cboTest)))
    
    Call AssertTrue("ControlTypeConsistency - CheckBox typename supported", _
                    Control_IsSupportedType(TypeName(frmTestControls.cbTest)))
End Sub


