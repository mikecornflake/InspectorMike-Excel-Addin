VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmXLLaunch 
   Caption         =   "Add / Edit Event"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "frmXLLaunch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmXLLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mButtonHandlers As Collection

Private Sub UserForm_Initialize()
    SetupForm
End Sub

Private Sub UserForm_Activate()
    CentreFormOverExcel Me
End Sub

Public Sub SetupForm()
    Set mButtonHandlers = New Collection
    
    ClearDynamicButtons
    BuildEventButtons
    LayoutForm
End Sub

Private Sub ClearDynamicButtons()
    Dim i As Long
    
    For i = fraHost.Controls.Count - 1 To 0 Step -1
        fraHost.Controls.Remove fraHost.Controls(i).Name
    Next i
End Sub

Private Sub BuildEventButtons()
    Const SHEET_FORMS As String = "xe.forms"
    Const SCROLLBAR_WIDTH As Single = 18
        
    Dim wsForms As Worksheet
    Dim colFormID As Long
    Dim colCaption As Long
    Dim colType As Long
    Dim lastRow As Long
    Dim iRow As Long
    
    Dim sFormID As String
    Dim sCaption As String
    Dim sType As String
    
    Dim btn As MSForms.CommandButton
    Dim handler As clsLaunchButtonHandler
    
    Dim curTop As Single
    Dim margin As Single
    Dim buttonHeight As Single
    Dim buttonWidth As Single
    
    If Not WorksheetExists(SHEET_FORMS) Then
        MsgBox "xe.forms not found.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    Set wsForms = ActiveWorkbook.Worksheets(SHEET_FORMS)
    
    colFormID = FindColumnInSheet(wsForms, "FormID")
    colCaption = FindColumnInSheet(wsForms, "Caption")
    colType = FindColumnInSheet(wsForms, "Type")
    
    If (colFormID <= 0) Or (colCaption <= 0) Or (colType <= 0) Then
        MsgBox "xe.forms is missing required columns.", vbExclamation, "xlEventing"
        Exit Sub
    End If
    
    margin = 12
    buttonHeight = 28
    buttonWidth = fraHost.InsideWidth - (2 * margin) - SCROLLBAR_WIDTH - 4
    
    curTop = margin
    
    lastRow = LastUsedRow(wsForms)
    
    For iRow = 2 To lastRow
        sFormID = Trim$(CStr(wsForms.Cells(iRow, colFormID).Value))
        sCaption = Trim$(CStr(wsForms.Cells(iRow, colCaption).Value))
        sType = LCase$(Trim$(CStr(wsForms.Cells(iRow, colType).Value)))
        
        If sType = "event" Then
            Set btn = fraHost.Controls.Add("Forms.CommandButton.1", "btn_" & SafeControlSuffix(sFormID, iRow), True)
            btn.Caption = sCaption
            btn.Tag = sFormID
            btn.Left = margin
            btn.Top = curTop
            btn.Width = buttonWidth
            btn.Height = buttonHeight
            
            Set handler = New clsLaunchButtonHandler
            handler.Init btn, sFormID, Me
            mButtonHandlers.Add handler
            
            curTop = curTop + buttonHeight + 6
        End If
    Next iRow
    
    fraHost.ScrollBars = fmScrollBarsVertical
    fraHost.KeepScrollBarsVisible = fmScrollBarsVertical
    fraHost.ScrollHeight = curTop + margin
End Sub

Private Sub LayoutForm()
    Dim margin As Single
    
    margin = 12
    
    fraHost.Left = margin
    fraHost.Top = margin
    fraHost.Width = Me.InsideWidth - (2 * margin)
    fraHost.Height = Me.InsideHeight - (2 * margin)
End Sub

Private Sub UserForm_Resize()
    LayoutForm
End Sub

Public Sub LaunchForm(ByVal pFormID As String)
    Unload Me
    ShowXlEventingForm pFormID, -1
End Sub
