Attribute VB_Name = "Addin"

'Callback for btnAbout onAction
Sub About_Callback(control As IRibbonControl)
    Call ShowAboutHtml
End Sub

'Callback for btnTidyEventExport onAction
Sub Tidy_Nexus_Event_Export_Callback(control As IRibbonControl)
    Tidy_Event_Export
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
    BasicTidy
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

Sub ConvertSelectedToTitleCase_Callback(control As IRibbonControl)
    ConvertSelectedToTitleCase
End Sub

Sub ConvertSelectedToSentenceCase_Callback(control As IRibbonControl)
    ConvertSelectedToSentenceCase
End Sub

Sub ConvertSelectedToUpperCase_Callback(control As IRibbonControl)
    ConvertSelectedToUpperCase
End Sub

Sub ConvertSelectedToLowerCase_Callback(control As IRibbonControl)
    ConvertSelectedToLowerCase
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

Sub Show_Eventing_Admin_Callback(control As IRibbonControl)
    Call ShowXLAdminForm
End Sub

Sub Show_Eventing_Launch_Callback(control As IRibbonControl)
    ShowXlLaunchForm
End Sub

Public Sub Show_Eventing_Edit_Callback(control As IRibbonControl)
    ShowXlEventingForm_EditOrAppendFromActiveSheet
End Sub

Sub ShowAboutHtml()
    Dim tempPath As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim html As String

    tempPath = Environ("TEMP")
    filePath = tempPath & "\about.html"

    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html><head><meta charset='UTF-8'><title>About</title>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body { font-family: Segoe UI, sans-serif; margin: 2em; background: #f9f9f9; color: #333; }" & vbCrLf
    html = html & "h1, h2 { color: #2a4d7c; } ul { margin-left: 1em; } .date { font-weight: bold; }" & vbCrLf
    html = html & "</style></head><body>" & vbCrLf

    html = html & "<h1>Inspector Mike 2.0 Excel Addin</h1>" & vbCrLf
    html = html & "<p><strong>Last updated:</strong> 16 August 2025</p>" & vbCrLf
    html = html & "<p><strong>Author:</strong> Mike Thompson (<a href='mailto:mike.cornflake@gmail.com'>mike.cornflake@gmail.com</a>)</p>" & vbCrLf

    html = html & "<h2>Contributions</h2>" & vbCrLf
    html = html & "<ul>" & vbCrLf
    html = html & " <li>Chris Merrick (2004)</li>" & vbCrLf
    html = html & " <li>MSDN, <a href='https://stackoverflow.com/' target='_blank' rel='noopener noreferrer'>StackOverflow</a> (attributes in code)</li>" & vbCrLf
    html = html & " <li>Ion Cristian Buse (2012): <a href='https://github.com/cristianbuse/VBA-FileTools' target='_blank' rel='noopener noreferrer'>VBA FileTools</a> (OneDrive/SharePoint support)</li>" & vbCrLf
    html = html & " <li><a href='https://copilot.microsoft.com' target='_blank' rel='noopener noreferrer'>Copilot/ChatGPT-5</a> (2025+): Unit Test framework, code review & inline documentation</li>" & vbCrLf
    html = html & "</ul>" & vbCrLf

    html = html & "<h2>Recent Changes</h2>" & vbCrLf
    html = html & "<ul>" & vbCrLf
    
    html = html & " <li><span class='date'>20/04/2026:</span> </li>" & vbCrLf
    html = html & "  <ul>" & vbCrLf
    html = html & "   <li>Draft Excel Eventing</li>" & vbCrLf
    html = html & "   <li>Extended Sorts to handle Collection of Array & extended unit tests</li>" & vbCrLf
    html = html & "   <li>ChatGPT 5.3 pointed out my WorkSheet Calls are not safe.  Need to call <worksheet>.Cells explicitly, I've been calling <implied activesheet>.cells.  Started making these changes</li>" & vbCrLf
    html = html & "  </ul>" & vbCrLf
    html = html & " </li>" & vbCrLf
    
    html = html & " <li><span class='date'>16/08/2025:</span> </li>" & vbCrLf
    html = html & "  <ul>" & vbCrLf
    html = html & "   <li>First contributions by copilot/ChatGPT-5</li>" & vbCrLf
    html = html & "   <li>Added Unit Test framework, and started added tests to units</li>" & vbCrLf
    html = html & "   <li>Re-branded tools from DOF to Inspector Mike</li>" & vbCrLf
    html = html & "   <li>Merged standalone code supporting Integrity Elements (ongoing)</li>" & vbCrLf
    html = html & "   <li>Imported Sense.Structures Export Processing routines developed 06/08 to 12/08/2025</li>" & vbCrLf
    html = html & "   <li>Removed Sharepoint/OneDrive support (retaining VBA-FileTools for other functionality)</li>" & vbCrLf
    html = html & "   <li>Removed contributions from DOF personnel where I forgot to get permission to continue using (file listing utilities)</li>" & vbCrLf
    html = html & "   <li>Removed DOF Project Specific processing</li>" & vbCrLf
    html = html & "   <li>Removed NIS Talisman and Maersk processing routines</li>" & vbCrLf
    html = html & "   <li>Ported About to HTML</li>" & vbCrLf
    html = html & "  </ul>" & vbCrLf
    html = html & " </li>" & vbCrLf
    html = html & "</ul>" & vbCrLf
    
    html = html & "<p><strong>Recommended install location:</strong> <code>%appdata%\Microsoft\AddIns</code></p>" & vbCrLf
    
    html = html & "<h2>TODO</h2>" & vbCrLf
    html = html & "<ul>" & vbCrLf
    html = html & " <li>Database routines have been entirely worked over, need testing.  I'm concerned about the removal of multiple <b>On Error Resume Next</b></li>" & vbCrLf
    html = html & " <li>Complete code review and Unit Tests</li>" & vbCrLf
    html = html & " <li>Complete Integrity Elements merging - add Ribbon Entries</li>" & vbCrLf
    html = html & " <li>Sanity check Nexus 6 Event Export processing.  Look for DOF IP and remove/back room reengineer.</li>" & vbCrLf
    html = html & " <li>Add File Import routines (back room engineer DOF routines)</li>" & vbCrLf
    html = html & " <li>Continue Interpolation routines</li>" & vbCrLf
    html = html & " <li>Add UI for Column Formatting.  Persist settings and include option to export/import</li>" & vbCrLf
    html = html & " <li>Looks like Column Renaming is specific to Nexus 6 - move into Nexus tab</li>" & vbCrLf
    html = html & "</ul>" & vbCrLf
    
    html = html & "<h2>History</h2>" & vbCrLf
    html = html & "<ul>" & vbCrLf
    html = html & " <li><span class='date'>2025+:</span> Inspector Mike 2.0 Pty Ltd</li>" & vbCrLf
    html = html & " <ul>" & vbCrLf
    html = html & "  <details>" & vbCrLf
    html = html & "   <li>The journey continues...</li>" & vbCrLf
    html = html & "   <li><span class='date'>16/08/2025:</span> </li>" & vbCrLf
    html = html & "    <ul>" & vbCrLf
    html = html & "     <li>First contributions by copilot/ChatGPT-5</li>" & vbCrLf
    html = html & "     <li>Added Unit Test framework, and started added tests to units</li>" & vbCrLf
    html = html & "     <li>Re-branded tools from DOF to Inspector Mike</li>" & vbCrLf
    html = html & "     <li>Merged standalone code supporting Integrity Elements (ongoing)</li>" & vbCrLf
    html = html & "     <li>Imported Sense.Structures Export Processing routines developed 06/08 to 12/08/2025</li>" & vbCrLf
    html = html & "     <li>Removed Sharepoint/OneDrive support (retaining VBA-FileTools for other functionality)</li>" & vbCrLf
    html = html & "     <li>Removed contributions from DOF personnel where I forgot to get permission to continue using (file listing utilities)</li>" & vbCrLf
    html = html & "     <li>Removed DOF Project Specific processing</li>" & vbCrLf
    html = html & "     <li>Removed NIS Talisman and Maersk processing routines</li>" & vbCrLf
    html = html & "     <li>Ported About to HTML</li>" & vbCrLf
    html = html & "    </ul>" & vbCrLf
    html = html & "   </li>" & vbCrLf
    html = html & "  </details>" & vbCrLf
    html = html & " </ul>" & vbCrLf
    html = html & " <li><span class='date'>2014 to 2022:</span> DOF Subsea Pty Ltd (Perth)</li>" & vbCrLf
    html = html & " <ul>" & vbCrLf
    html = html & "  <details>" & vbCrLf
    html = html & "   <li><span class='date'>06/06/2024:</span> Updates for 2023 AUV Re-reprocessing</li>" & vbCrLf
    html = html & "   <li><span class='date'>11/04/2024:</span> Updated LibraryFileTools to latest VBA-FileTools</li>" & vbCrLf
    html = html & "   <li><span class='date'>09/04/2024:</span> Updates to 2023 AUV Routines</li>" & vbCrLf
    html = html & "   <li><span class='date'>19/12/2023:</span> Minor update (format Clock Column)</li>" & vbCrLf
    html = html & "   <li><span class='date'>09/11/2023:</span> Updates to 2023 AUV Routines (to 22/11/2023)</li>" & vbCrLf
    html = html & "   <li><span class='date'>06/10/2023:</span> Commenced IntegrityElements Support</li>" & vbCrLf
    html = html & "   <li><span class='date'>27/09/2023:</span> Commenced LibraryGantt</li>" & vbCrLf
    html = html & "   <li><span class='date'>01/09/2023:</span> Added routines for 2023 ABU AUV Processing</li>" & vbCrLf
    html = html & "   <li><span class='date'>30/08/2023:</span> Removed conflicts between OneDrive/Sharepoint routines and existing (default to OneDrive version)</li>" & vbCrLf
    html = html & "   <li><span class='date'>30/08/2023:</span> Added handler for No Finding ID in Nexus6 Event Export</li>" & vbCrLf
    html = html & "   <li><span class='date'>22/08/2023:</span> Added LibraryDate, SaveAndBackup, LibraryFileTools and limited support for OneDrive/Sharepoint</li>" & vbCrLf
    html = html & "   <li><span class='date'>12/06/2023:</span> Minor fixes to ADO connection</li>" & vbCrLf
    html = html & "   <li><span class='date'>09/05/2023:</span> Fixed Excel 2019 compatibility (no formula2)</li>" & vbCrLf
    html = html & "   <li><span class='date'>03/05/2023:</span> Minor fixes to ADO connection</li>" & vbCrLf
    html = html & "   <li><span class='date'>16/02/2023:</span> Added Support for Nexus 6 event export (Structure)</li>" & vbCrLf
    html = html & "   <li><span class='date'>09/01/2023:</span> Initial Support for Nexus 6 event export (Pipeline)</li>" & vbCrLf
    html = html & "   <li><span class='date'>07/12/2022:</span> Added initial Compare Sheets support</li>" & vbCrLf
    html = html & "   <li><span class='date'>31/08/2022:</span> Unified workflow for Nexus 6/Malampaya Survey Processing</li>" & vbCrLf
    html = html & "   <li><span class='date'>26/08/2022:</span> Improved RibbonUI - including adding String Case Formatting</li>" & vbCrLf
    html = html & "   <li><span class='date'>18/08/2022:</span> Additions for Nexus 6/Malampaya Survey Processing</li>" & vbCrLf
    html = html & "   <li><span class='date'>07/12/2021:</span> Fixed & improved run time for List Files In Folder</li>" & vbCrLf
    html = html & "   <li><span class='date'>19/10/2020:</span> Added Microsoft Planner Export tidy</li>" & vbCrLf
    html = html & "   <li><span class='date'>10/10/2019:</span> Added draft support for Conditional Formatting (no menu)</li>" & vbCrLf
    html = html & "   <li><span class='date'>07/01/2019:</span> Nexus 5 Export: Ignore files that are reporting an off 'File Not Found'. These will need to be manually processed.</li>" & vbCrLf
    html = html & "   <li><span class='date'>06/06/2018:</span> Expanded Nexus Export Multimedia renaming to handle files not being moved</li>" & vbCrLf
    html = html & "   <li><span class='date'>26/04/2018:</span> Added String Case handling (TODO - Expand RibbonBar)</li>" & vbCrLf
    html = html & "   <li><span class='date'>19/04/2018:</span> Added ListFilesInFolder handling (TODO - Expand RibbonBar)</li>" & vbCrLf
    html = html & "   <li><span class='date'>03/03/2018:</span> Remove all Conditional Formatting from Nexus Exports</li>" & vbCrLf
    html = html & "   <li><span class='date'>03/03/2018:</span> Stop crashing when no Findings exist</li>" & vbCrLf
    html = html & "   <li><span class='date'>12/02/2018:</span> Optionally Rename and Move all multimedia in Nexus Exports</li>" & vbCrLf
    html = html & "   <li><span class='date'>14/12/2017:</span> Added DOF branding as a courtesy</li>" & vbCrLf
    html = html & "   <li><span class='date'>14/12/2017:</span> Added Ribbon, installers and initial documentation</li>" & vbCrLf
    html = html & "  </details>" & vbCrLf
    html = html & " </ul>" & vbCrLf
    html = html & " <li><span class='date'>2007 to 2014:</span> Inspector Mike Pty Ltd</li>" & vbCrLf
    html = html & " <ul>" & vbCrLf
    html = html & "  <details>" & vbCrLf
    html = html & "   <li>Ongoing Event Exporter enhancements for Talisman and Maersk Denmark</li>" & vbCrLf
    html = html & "   <li>Ongoing Maintenace and usage</li>" & vbCrLf
    html = html & "  </details>" & vbCrLf
    html = html & " </ul>" & vbCrLf
    html = html & " <li><span class='date'>2004 to 2007:</span> Netlink Inspection Services (NIS)</li>" & vbCrLf
    html = html & " <ul>" & vbCrLf
    html = html & "  <details>" & vbCrLf
    html = html & "   <li>2005 to 2007: Multiple library enhancements to support offshore operations with assorted Netlink modules</li>" & vbCrLf
    html = html & "   <li>06/2005: Event Exporter enhancements for Talisman Malaysia & Maersk Denmark operations and reporting</li>" & vbCrLf
    html = html & "   <li>04/2004: Event Exporter enhancements for Nexen and Maersk Oil UK reportingk</li>" & vbCrLf
    html = html & "   <li>02/2004: Initial Event Export Processor in support of Shell Malampaya</li>" & vbCrLf
    html = html & "   <li>04/2004: Created initial library to support video conversion and log imports for China National Oil Company</li>" & vbCrLf
    html = html & "  </details>" & vbCrLf
    html = html & " </ul>" & vbCrLf
    html = html & "</ul>" & vbCrLf
    html = html & "</body></html>"

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, html
    Close #fileNum

    ActiveWorkbook.FollowHyperlink filePath
End Sub

