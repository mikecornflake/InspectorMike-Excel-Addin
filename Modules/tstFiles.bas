Attribute VB_Name = "tstFiles"
' 2025 08 16 - Code review by Copilot/ChatGPT-5 and Mike T

Option Explicit
Option Private Module

Public Sub Test_LibraryFiles()
    ActiveTestModule = "libFiles"

    ' Path_AddTrailingDelimiter
    Call AssertEqual("Path_AddTrailingDelimiter - already has slash", "C:\Test\", Path_AddTrailingDelimiter("C:\Test\"))
    Call AssertEqual("Path_AddTrailingDelimiter - missing slash", "C:\Test\", Path_AddTrailingDelimiter("C:\Test"))

    ' Path_RemoveTrailingDelimiter
    Call AssertEqual("Path_RemoveTrailingDelimiter - has slash", "C:\Test", Path_RemoveTrailingDelimiter("C:\Test\"))
    Call AssertEqual("Path_RemoveTrailingDelimiter - no slash", "C:\Test", Path_RemoveTrailingDelimiter("C:\Test"))

    ' Path_GetFolder
    Call AssertEqual("Path_GetFolder - full path", "C:\Test", Path_GetFolder("C:\Test\MyFile.xlsx"))
    Call AssertEqual("Path_GetFolder - no slashes", "", Path_GetFolder("MyFile.xlsx"))

    ' Path_GetFileName
    Call AssertEqual("Path_GetFileName - full path", "MyFile.xlsx", Path_GetFileName("C:\Test\MyFile.xlsx"))

    ' Path_GetFileNameNoExt
    Call AssertEqual("Path_GetFileNameNoExt - with extension", "MyFile", Path_GetFileNameNoExt("C:\Test\MyFile.xlsx"))
    Call AssertEqual("Path_GetFileNameNoExt - no extension", "MyFile", Path_GetFileNameNoExt("C:\Test\MyFile"))

    ' Path_GetExtension
    Call AssertEqual("Path_GetExtension - with extension", ".xlsx", Path_GetExtension("C:\Test\MyFile.xlsx"))
    Call AssertEqual("Path_GetExtension - no extension", "", Path_GetExtension("C:\Test\MyFile"))

    ' File_SanitizeName
    Call AssertEqual("File_SanitizeName - invalid chars", "My-File-Name", File_SanitizeName("My/File:Name", "-"))
    Call AssertEqual("File_SanitizeName - custom replacement", "My_File_Name", File_SanitizeName("My/File:Name", "_"))

    ' ActiveWorkbookPath (indirect test)
    Call AssertTrue("ActiveWorkbookPath ends with slash", Right(ActiveWorkbookPath, 1) = "\")

    ' File_EnsureFolder (integration-style test using %TEMP%)
    Dim tempPath As String
    Dim testPath As String
    
    tempPath = Environ("TEMP")
    testPath = Path_AddTrailingDelimiter(tempPath) & "TestFolder1\TestFolder2\TestFolder3"
    
    Call File_EnsureFolder(testPath)
    Call AssertTrue("File_EnsureFolder - folder created in TEMP", Folder_Exists(testPath))

    ' File_List (basic test)
    Dim col As Collection
    Set col = File_List(tempPath, "*.xls*")
    Call AssertTrue("File_List - returns collection", TypeName(col) = "Collection")
End Sub
