Attribute VB_Name = "libWorkbook"
' 18 May 2026
'  Rationalised namespace and refacted from libExcel

Option Explicit
Option Private Module

' Returns the full path of the active workbook
Public Function Workbook_ActiveLocalFilename() As String
    Workbook_ActiveLocalFilename = ActiveWorkbook.FullName
End Function

' Returns the directory of the active workbook with trailing backslash
Public Function Workbook_ActiveDirectory() As String
    Workbook_ActiveDirectory = Path_AddTrailingDelimiter(Path_GetFolder(Workbook_ActiveLocalFilename))
End Function

' Returns the active workbook's folder path with trailing delimiter
Public Function Workbook_ActivePath() As String
    Workbook_ActivePath = Workbook_ActiveDirectory
End Function

' Saves a backup copy of the active workbook in an "Archive" subfolder
Public Sub Workbook_SaveAndBackup()
    Dim sOriginalFile As String, sOriginalFolder As String
    Dim sNewFile As String, sBackupDir As String, sTemp As String
    Dim fso As Object

    On Error GoTo ErrHandler

    sTemp = ActiveWorkbook.FullName
    sOriginalFile = ActiveWorkbook.Name

    ' Check if the file has been saved yet
    If sTemp = sOriginalFile Then
        Application_Original_Save_As_Dialog
        sTemp = Workbook_ActiveLocalFilename
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

' Saves the active sheet as XLSX
Public Sub Workbook_SaveAsXLSX()
    Dim sExt As String, sFilename As String

    On Error GoTo ErrHandler

    sFilename = Workbook_ActiveLocalFilename
    sExt = LCase(Path_GetExtension(sFilename))

    If sFilename = "" Then
        Application_Original_Save_As_Dialog
    Else
        If sExt = "" Then
            sFilename = sFilename & ".xlsx"
        Else
            sFilename = Text_Replace(sFilename, sExt, ".xlsx")
        End If
        ActiveSheet.SaveAs filename:=sFilename, FileFormat:=xlWorkbookDefault, Local:=True
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error during SaveAsXLSX: " & Err.Description, vbExclamation
End Sub

Public Function Workbook_DuplicateActive(ANewFullname As String) As Workbook
    Dim oBook As Workbook
    Dim sNewPath As String, sNewImageDir As String
    Dim bHasImages As Boolean
    Dim sSourcePath As String, sSourceFilename As String, sSourceImageDir As String
    Dim oSheet As Worksheet
    Dim sFilename As String
    
    
    sFilename = Workbook_ActiveLocalFilename
    sSourcePath = Workbook_ActivePath
    sSourceFilename = Path_GetFileNameNoExt(sFilename)
    sSourceImageDir = sSourceFilename + "_Images"
    
    sNewPath = Path_AddTrailingDelimiter(Path_GetFolder(ANewFullname))
    sNewImageDir = Path_GetFileNameNoExt(ANewFullname) + "_Images"
    
    bHasImages = Folder_Exists(sSourcePath + sSourceImageDir)
    
    File_EnsureFolder (sNewPath)
    ActiveWorkbook.SaveCopyAs (ANewFullname)
    
    Set oBook = Workbooks.Open(ANewFullname)
    
    If bHasImages Then
        Application.StatusBar = "Copying files from " + sSourcePath + sSourceImageDir + ".  This may take a while."
        Call CopyFolder(sSourcePath + sSourceImageDir, sNewPath + sNewImageDir)
        
        'Iterate over each worksheet and correct all hyperlinks
        Application.StatusBar = "Correcting hyperlinks"
        If oBook.Worksheets.Count > 0 Then
            For Each oSheet In oBook.Worksheets
                DoEvents
                oSheet.Activate
                
                oSheet.Cells.Select
                Selection.Replace What:=sSourceImageDir, Replacement:=sNewImageDir, LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                    
                oSheet.Cells(2, 1).Select
            Next oSheet
            
            oBook.Worksheets(1).Activate
        End If
    End If
    
    oBook.Save
    
    Application.StatusBar = ""
    Set Workbook_DuplicateActive = oBook
End Function
