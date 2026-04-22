Attribute VB_Name = "libFiles"
' 2025 08 16 - Code review by Copilot/ChatGPT-5 and Mike T
'            - Stopped using OneDrive/Sharepoint URL safe code.  MS Broke the API too many times

Option Explicit

Public Sub Test_LibraryFiles()
    ActiveTestModule = "libFiles"

    ' AddTrailingDelimiter
    Call AssertEqual("AddTrailingDelimiter - already has slash", "C:\Test\", AddTrailingDelimiter("C:\Test\"))
    Call AssertEqual("AddTrailingDelimiter - missing slash", "C:\Test\", AddTrailingDelimiter("C:\Test"))

    ' RemoveTrailingDelimiter
    Call AssertEqual("RemoveTrailingDelimiter - has slash", "C:\Test", RemoveTrailingDelimiter("C:\Test\"))
    Call AssertEqual("RemoveTrailingDelimiter - no slash", "C:\Test", RemoveTrailingDelimiter("C:\Test"))

    ' ExtractFolder
    Call AssertEqual("ExtractFolder - full path", "C:\Test", ExtractFolder("C:\Test\MyFile.xlsx"))
    Call AssertEqual("ExtractFolder - no slashes", "", ExtractFolder("MyFile.xlsx"))

    ' ExtractFilename
    Call AssertEqual("ExtractFilename - full path", "MyFile.xlsx", ExtractFilename("C:\Test\MyFile.xlsx"))

    ' ExtractFilenameOnly
    Call AssertEqual("ExtractFilenameOnly - with extension", "MyFile", ExtractFilenameOnly("C:\Test\MyFile.xlsx"))
    Call AssertEqual("ExtractFilenameOnly - no extension", "MyFile", ExtractFilenameOnly("C:\Test\MyFile"))

    ' ExtractFilenameExt
    Call AssertEqual("ExtractFilenameExt - with extension", ".xlsx", ExtractFilenameExt("C:\Test\MyFile.xlsx"))
    Call AssertEqual("ExtractFilenameExt - no extension", "", ExtractFilenameExt("C:\Test\MyFile"))

    ' ValidateFilename
    Call AssertEqual("ValidateFilename - invalid chars", "My-File-Name", ValidateFilename("My/File:Name", "-"))
    Call AssertEqual("ValidateFilename - custom replacement", "My_File_Name", ValidateFilename("My/File:Name", "_"))

    ' ActiveWorkbookPath (indirect test)
    Call AssertTrue("ActiveWorkbookPath ends with slash", Right(ActiveWorkbookPath, 1) = "\")

    ' ForceDirectories (integration-style test using %TEMP%)
    Dim tempPath As String
    Dim testPath As String
    
    tempPath = Environ("TEMP")
    testPath = AddTrailingDelimiter(tempPath) & "TestFolder1\TestFolder2\TestFolder3"
    
    Call ForceDirectories(testPath)
    Call AssertTrue("ForceDirectories - folder created in TEMP", FolderExists(testPath))

    ' FilesInFolder (basic test)
    Dim col As Collection
    Set col = FilesInFolder(tempPath, "*.xls*")
    Call AssertTrue("FilesInFolder - returns collection", TypeName(col) = "Collection")
End Sub

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
    ActiveWorkbookDirectory = AddTrailingDelimiter(ExtractFolder(ActiveWorkbookLocalFilename))
End Function

' For use when bulk converting PDFs
Public Sub SaveAsPDF()
    Dim sFilename As String
    
    sFilename = SwapString(ActiveWorkbookLocalFilename, ".xlsx", ".pdf")
    sFilename = SwapString(sFilename, ".xls", ".pdf")
    sFilename = SwapString(sFilename, ".pdf", ActiveSheet.Name & ".pdf")
    
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
    sBackupDir = AddTrailingDelimiter(sOriginalFolder) & "Archive"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(sBackupDir) Then fso.CreateFolder sBackupDir

    sNewFile = Format(Now, "yyyy mm dd-hh mm ss") & " - " & sOriginalFile

    ' Save backup copy
    ActiveWorkbook.SaveCopyAs filename:=AddTrailingDelimiter(sBackupDir) & sNewFile

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
    Dim i As Integer

    On Error GoTo ErrHandler

    sFolder = AddTrailingDelimiter(ExtractFolder(AFilename))
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
    sFilenameOnly = ExtractFilenameOnly(sFilename)
    sExt = LCase(ExtractFilenameExt(sFilename))
    sSheetname = ActiveSheet.Name

    If AFilename <> "" Then
        sCSV = ActiveWorkbookDirectory & AFilename
    Else
        If (sExt = ".xls") Or (sExt = ".xlsx") Then
            If AIncludeSheetName Then
                sCSV = SwapString(sFilename, sExt, " - " & sSheetname & ".csv")
            Else
                sCSV = SwapString(sFilename, sExt, ".csv")
            End If
        ElseIf sFilenameOnly = sSheetname Then
            sCSV = SwapString(sFilename, sExt, ".csv")
        Else
            sCSV = SwapString(sFilename, ActiveWorkbook.Name, sSheetname & ".csv")
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
    sExt = LCase(ExtractFilenameExt(sFilename))

    If sFilename = "" Then
        Original_Save_As_Dialog
    Else
        If sExt = "" Then
            sFilename = sFilename & ".csv"
        Else
            sFilename = SwapString(sFilename, sExt, ".xlsx")
        End If
        ActiveSheet.SaveAs filename:=sFilename, FileFormat:=xlWorkbookDefault, Local:=True
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error during SaveAsXLSX: " & Err.Description, vbExclamation
End Sub

' Adds a trailing backslash to a folder path
Public Function AddTrailingDelimiter(sInput As String) As String
    If Right(sInput, 1) <> "\" Then
        AddTrailingDelimiter = sInput & "\"
    Else
        AddTrailingDelimiter = sInput
    End If
End Function

' Removes a trailing backslash from a folder path
Public Function RemoveTrailingDelimiter(sInput As String) As String
    If Right(sInput, 1) = "\" Then
        RemoveTrailingDelimiter = Left(sInput, Len(sInput) - 1)
    Else
        RemoveTrailingDelimiter = sInput
    End If
End Function

' Extracts the folder path from a full file path
Public Function ExtractFolder(sInput As String) As String
    If InStr(sInput, "\") Then
        ExtractFolder = Left(sInput, InStrRev(sInput, "\") - 1)
    Else
        ExtractFolder = ""
    End If
End Function

' Extracts the filename from a full path
Public Function ExtractFilename(sInput As String) As String
    Dim fso As Object
    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    ExtractFilename = fso.GetFileName(sInput)
    Set fso = Nothing
    Exit Function

ErrHandler:
    ExtractFilename = ""
End Function

' Extracts the filename without extension
Public Function ExtractFilenameOnly(sInput As String) As String
    Dim sTemp As String, iPos As Long
    sTemp = ExtractFilename(sInput)
    iPos = InStrRev(sTemp, ".")
    If iPos > 0 Then
        ExtractFilenameOnly = Left(sTemp, iPos - 1)
    Else
        ExtractFilenameOnly = sTemp
    End If
End Function

' Extracts the file extension
Public Function ExtractFilenameExt(sInput As String) As String
    Dim iPos As Integer
    iPos = InStrRev(sInput, ".")
    If iPos > 0 Then
        ExtractFilenameExt = Mid(sInput, iPos)
    Else
        ExtractFilenameExt = ""
    End If
End Function

' Returns the active workbook's folder path with trailing delimiter
Public Function ActiveWorkbookPath() As String
    ActiveWorkbookPath = AddTrailingDelimiter(ExtractFolder(ActiveWorkbookLocalFilename))
End Function

' Creates all folders in a path if they don't exist
Public Sub ForceDirectories(sFolderPath As String)
    Dim x As Variant, i As Integer, strPath As String
    x = Split(sFolderPath, "\")
    For i = 0 To UBound(x)
        strPath = strPath & x(i) & "\"
        If Not FolderExists(strPath) Then MkDir strPath
    Next i
End Sub

' Replaces invalid characters in a filename with a safe replacement character
Function ValidateFilename(sFilename As String, Optional sReplacement As String = "-") As String
    Dim sTemp As String

    On Error GoTo ErrHandler

    sTemp = sFilename

    ' Replace characters that are not allowed in Windows filenames
    sTemp = Replace(sTemp, "\", sReplacement)
    sTemp = Replace(sTemp, "/", sReplacement)
    sTemp = Replace(sTemp, ":", sReplacement)
    sTemp = Replace(sTemp, """", sReplacement)
    sTemp = Replace(sTemp, "*", sReplacement)
    sTemp = Replace(sTemp, "?", sReplacement)
    sTemp = Replace(sTemp, "<", sReplacement)
    sTemp = Replace(sTemp, ">", sReplacement)
    sTemp = Replace(sTemp, "|", sReplacement)

    ValidateFilename = sTemp
    Exit Function

ErrHandler:
    ' If something goes wrong, return a generic safe filename
    ValidateFilename = "invalid_filename"
End Function

' Returns a collection of file paths in the specified folder that match the given file mask
Public Function FilesInFolder(AFolder As String, AFileMask As String) As Collection
    Dim folderPath As String
    Dim filename As String
    Dim txtFilesCollection As New Collection

    On Error GoTo ErrHandler

    ' Ensure folder path ends with a backslash
    folderPath = AFolder
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Get the first matching file
    filename = Dir(folderPath & AFileMask)

    ' Loop through all matching files
    Do While filename <> ""
        txtFilesCollection.Add folderPath & filename
        filename = Dir
    Loop

    ' Return the collection
    Set FilesInFolder = txtFilesCollection
    Exit Function

ErrHandler:
    MsgBox "Error in FilesInFolder: " & Err.Description, vbExclamation
    Set FilesInFolder = Nothing
End Function

' Wrappers...  I keep forgetting the new names....
Function FolderExists(AFolder As String)
    FolderExists = IsFolder(AFolder)
End Function

Function FileExists(AFile As String)
    FileExists = IsFile(AFile)
End Function

