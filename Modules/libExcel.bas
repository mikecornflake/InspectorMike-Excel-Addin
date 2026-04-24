Attribute VB_Name = "libExcel"
' 24 Apr 2026 - refactored from libFiles
'            - Stopped using OneDrive/Sharepoint URL safe code.  MS Broke the API too many times

' Opens the standard Save As dialog
Public Sub Original_Save_As_Dialog()
    Application.Dialogs(xlDialogSaveAs).Show
End Sub

' Returns the full path of the active workbook
Public Function ActiveWorkbookLocalFilename() As String
    ActiveWorkbookLocalFilename = ActiveWorkbook.FullName
End Function

' Returns the directory of the active workbook with trailing backslash
Public Function ActiveWorkbookDirectory() As String
    ActiveWorkbookDirectory = Path_AddTrailingDelimiter(Path_GetFolder(ActiveWorkbookLocalFilename))
End Function

' Returns the active workbook's folder path with trailing delimiter
Public Function ActiveWorkbookPath() As String
    ActiveWorkbookPath = Path_AddTrailingDelimiter(Path_GetFolder(ActiveWorkbookLocalFilename))
End Function

' For use when bulk converting PDFs
Public Sub SaveAsPDF()
    Dim sFilename As String
    
    sFilename = Text_Replace(ActiveWorkbookLocalFilename, ".xlsx", ".pdf")
    sFilename = Text_Replace(sFilename, ".xls", ".pdf")
    sFilename = Text_Replace(sFilename, ".pdf", ActiveSheet.Name & ".pdf")
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=sFilename, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

' Saves a backup copy of the active workbook in an "Archive" subfolder
Public Sub SaveAndBackup()
    Dim sOriginalFile As String, sOriginalFolder As String
    Dim sNewFile As String, sBackupDir As String, sTemp As String
    Dim fso As Object

    On Error GoTo ErrHandler

    sTemp = ActiveWorkbook.FullName
    sOriginalFile = ActiveWorkbook.Name

    ' Check if the file has been saved yet
    If sTemp = sOriginalFile Then
        Original_Save_As_Dialog
        sTemp = ActiveWorkbookLocalFilename
        sOriginalFile = ActiveWorkbook.Name
        If sTemp = sOriginalFile Then Exit Sub ' User cancelled
    End If

    sOriginalFolder = Mid(sTemp, 1, Len(sTemp) - Len(sOriginalFile))
    sBackupDir = Path_AddTrailingDelimiter(sOriginalFolder) & "Archive"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sBackupDir) Then fso.CreateFolder sBackupDir

    sNewFile = Format(Now, "yyyy mm dd-hh mm ss") & " - " & sOriginalFile

    ' Save backup copy
    ActiveWorkbook.SaveCopyAs filename:=Path_AddTrailingDelimiter(sBackupDir) & sNewFile

    ' Save original file
    ActiveWorkbook.Save
    Exit Sub

ErrHandler:
    MsgBox "Error during SaveAndBackup: " & Err.Description, vbExclamation
End Sub

' Exports the current worksheet to a new XLSX workbook
Public Sub ExportCurrentWorksheetAsXLSX(Optional AFilename As String = "")
    Dim sFilename As String, sFolder As String
    Dim oNew As Workbook, oCurrent As Workbook, oSheet As Worksheet
    Dim i As Long

    On Error GoTo ErrHandler

    sFolder = Path_AddTrailingDelimiter(Path_GetFolder(AFilename))
    If sFolder = "\" Then sFolder = ActiveWorkbookDirectory

    If AFilename <> "" Then
        sFilename = sFolder & AFilename
    Else
        sFilename = sFolder & ActiveSheet.Name
    End If

    Set oSheet = ActiveSheet
    Set oCurrent = ActiveWorkbook
    Set oNew = Workbooks.Add

    oSheet.Copy Before:=oNew.Sheets(1)

    ' Remove default sheets
    Application.DisplayAlerts = False
    For i = oNew.Sheets.Count To 2 Step -1
        oNew.Sheets(i).Delete
    Next i
    Application.DisplayAlerts = True

    oNew.SaveAs filename:=sFilename & ".xlsx"
    oNew.Close
    oCurrent.Activate
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "Error during ExportCurrentWorksheetAsXLSX: " & Err.Description, vbExclamation
End Sub

' Exports the current worksheet as a CSV file
Public Sub ExportCurrentWorkSheetAsCSV(Optional AIncludeSheetName As Boolean = True, Optional AFilename As String = "")
    Dim sFilename As String, sFilenameOnly As String
    Dim sCSV As String, sExt As String, sSheetname As String

    On Error GoTo ErrHandler

    sFilename = ActiveWorkbookLocalFilename
    sFilenameOnly = Path_GetFileNameNoExt(sFilename)
    sExt = LCase(Path_GetExtension(sFilename))
    sSheetname = ActiveSheet.Name

    If AFilename <> "" Then
        sCSV = ActiveWorkbookDirectory & AFilename
    Else
        If (sExt = ".xls") Or (sExt = ".xlsx") Then
            If AIncludeSheetName Then
                sCSV = Text_Replace(sFilename, sExt, " - " & sSheetname & ".csv")
            Else
                sCSV = Text_Replace(sFilename, sExt, ".csv")
            End If
        ElseIf sFilenameOnly = sSheetname Then
            sCSV = Text_Replace(sFilename, sExt, ".csv")
        Else
            sCSV = Text_Replace(sFilename, ActiveWorkbook.Name, sSheetname & ".csv")
        End If
    End If

    ActiveSheet.SaveAs filename:=sCSV, FileFormat:=xlCSV, Local:=True
    ActiveSheet.Name = sSheetname

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs filename:=sFilename, FileFormat:=xlWorkbookDefault, Local:=True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "Error during ExportCurrentWorkSheetAsCSV: " & Err.Description, vbExclamation
End Sub

' Saves the active sheet as XLSX
Public Sub SaveAsXLSX()
    Dim sExt As String, sFilename As String

    On Error GoTo ErrHandler

    sFilename = ActiveWorkbookLocalFilename
    sExt = LCase(Path_GetExtension(sFilename))

    If sFilename = "" Then
        Original_Save_As_Dialog
    Else
        If sExt = "" Then
            sFilename = sFilename & ".csv"
        Else
            sFilename = Text_Replace(sFilename, sExt, ".xlsx")
        End If
        ActiveSheet.SaveAs filename:=sFilename, FileFormat:=xlWorkbookDefault, Local:=True
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error during SaveAsXLSX: " & Err.Description, vbExclamation
End Sub

