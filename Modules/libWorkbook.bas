Attribute VB_Name = "libWorkbook"
Option Explicit

Public Function Duplicate_ActiveBook(AFullname As String) As Workbook
    Dim oBook As Workbook
    Dim sNewPath As String, sNewImageDir As String
    Dim bHasImages As Boolean
    Dim sSourcePath As String, sSourceFilename As String, sSourceImageDir As String
    Dim oSheet As Worksheet
    Dim sFilename As String
    
    
    sFilename = ActiveWorkbookLocalFilename
    sSourcePath = Path_AddTrailingDelimiter(Path_GetFolder(sFilename))
    sSourceFilename = Path_GetFileNameNoExt(sFilename)
    sSourceImageDir = sSourceFilename + "_Images"
    
    sNewPath = Path_AddTrailingDelimiter(Path_GetFolder(AFullname))
    sNewImageDir = Path_GetFileNameNoExt(AFullname) + "_Images"
    
    bHasImages = Folder_Exists(sSourcePath + sSourceImageDir)
    
    File_EnsureFolder (sNewPath)
    ActiveWorkbook.SaveCopyAs (AFullname)
    
    Set oBook = Workbooks.Open(AFullname)
    
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
    Set Duplicate_ActiveBook = oBook
End Function
