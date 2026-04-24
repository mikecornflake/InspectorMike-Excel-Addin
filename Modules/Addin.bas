Attribute VB_Name = "Addin"

'Callback for btnAbout onAction
Sub About_Callback(control As IRibbonControl)
    Call ShowHelp("about.html")
End Sub

'Callback for btnTidyEventExport onAction
Sub Tidy_Nexus_Event_Export_Callback(control As IRibbonControl)
    Nexus5_Tidy_Event_Export
End Sub

'Callback for btnTidyNexus6EventExport onAction
Sub Tidy_Nexus6_Event_Export_Callback(control As IRibbonControl)
    frmNexus6EventExport.Show
End Sub

'Callback for btnShowOptions onAction
Sub Process_VW_Event_Callback(control As IRibbonControl)
    ShowOptions
End Sub

'Callback for btnBasicTableTidy onAction
Sub Basic_Table_Tidy_Callback(control As IRibbonControl)
    Call BasicTidy(ActiveSheet)
End Sub

' Callback
Sub Interpolate_Nav_To_3_Sec_Callback(control As IRibbonControl)
    Interpolate_Nav_To_3_Sec
End Sub

' Callback
Sub Format_Standard_Columns_Callback(control As IRibbonControl)
    BasicTidyAndFormatColumns
    MsgBox ("Columns formatted.  If working with a CSV, please double check formatting is correct before saving.")
End Sub

' Callback
Sub Rename_Columns_Callback(control As IRibbonControl)
    frmRenameColumns.Show
End Sub

'Callback for btnExportAsCSV onAction
Sub Export_CSV_Callback(control As IRibbonControl)
    ExportCurrentWorkSheetAsCSV
End Sub

'Callback for btnSaveAsPDF onAction
Sub SaveAs_PDF_Callback(control As IRibbonControl)
    SaveAsPDF
End Sub

Sub Text_TitleCase_Selection_Callback(control As IRibbonControl)
    Text_TitleCase_Selection
End Sub

Sub Text_SentenceCase_Selection_Callback(control As IRibbonControl)
    Text_SentenceCase_Selection
End Sub

Sub Text_Upper_Selection_Callback(control As IRibbonControl)
    Text_Upper_Selection
End Sub

Sub Text_Lower_Selection_Callback(control As IRibbonControl)
    Text_Lower_Selection
End Sub

Sub SaveAndBackup_Callback(control As IRibbonControl)
    SaveAndBackup
End Sub

Sub OriginalSaveAs_Callback(control As IRibbonControl)
    Application.Dialogs(xlDialogSaveAs).Show
End Sub

Sub PrepareNexusImportFromCurrentSheet_Callback(control As IRibbonControl)
    PrepareNexusImportFromCurrentSheet
End Sub

Sub CompareSheets_Callback(control As IRibbonControl)
    frmCompareSheets.Show
End Sub

Sub Eventing_Admin_Callback(control As IRibbonControl)
    Call ShowXLAdminForm
End Sub

Sub Eventing_Launch_Callback(control As IRibbonControl)
    ShowXlLaunchForm
End Sub

Public Sub Eventing_Edit_Callback(control As IRibbonControl)
    ShowXlEventingForm_EditOrAppendFromActiveSheet
End Sub

Public Sub Eventing_Set_DateTime_Callback(control As IRibbonControl)
    IntelligentlyInsertDateTime
End Sub

Public Sub ShowHelpAbout(control As IRibbonControl)
    ShowHelp "about.html"
End Sub

