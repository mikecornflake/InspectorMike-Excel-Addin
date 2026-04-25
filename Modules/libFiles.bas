Attribute VB_Name = "libFiles"
Option Explicit

' Adds a trailing backslash to a folder path
Public Function Path_AddTrailingDelimiter(ByVal pPath As String) As String
    If Right(pPath, 1) <> "\" Then
        Path_AddTrailingDelimiter = pPath & "\"
    Else
        Path_AddTrailingDelimiter = pPath
    End If
End Function

' Removes a trailing backslash from a folder path
Public Function Path_RemoveTrailingDelimiter(ByVal pPath As String) As String
    If Right(pPath, 1) = "\" Then
        Path_RemoveTrailingDelimiter = Left(pPath, Len(pPath) - 1)
    Else
        Path_RemoveTrailingDelimiter = pPath
    End If
End Function

' Extracts the folder path from a full file path
Public Function Path_GetFolder(ByVal pPath As String) As String
    If InStr(pPath, "\") Then
        Path_GetFolder = Left(pPath, InStrRev(pPath, "\") - 1)
    Else
        Path_GetFolder = ""
    End If
End Function

' Extracts the filename from a full path
Public Function Path_GetFileName(ByVal pPath As String) As String
    Dim iPos As Long
    iPos = InStrRev(pPath, "\")
    
    If iPos > 0 Then
        Path_GetFileName = Mid$(pPath, iPos + 1)
    Else
        Path_GetFileName = pPath
    End If
End Function

' Extracts the filename without extension
Public Function Path_GetFileNameNoExt(ByVal pPath As String) As String
    Dim sTemp As String, iPos As Long
    sTemp = Path_GetFileName(pPath)
    iPos = InStrRev(sTemp, ".")
    If iPos > 0 Then
        Path_GetFileNameNoExt = Left(sTemp, iPos - 1)
    Else
        Path_GetFileNameNoExt = sTemp
    End If
End Function

' Extracts the file extension
Public Function Path_GetExtension(ByVal pPath As String) As String
    Dim iPos As Long
    iPos = InStrRev(pPath, ".")
    If iPos > 0 Then
        Path_GetExtension = Mid(pPath, iPos)
    Else
        Path_GetExtension = ""
    End If
End Function

' Creates all folders in a path if they don't exist
Public Sub File_EnsureFolder(ByVal pPath As String)
    Dim x As Variant, i As Long, strPath As String
    x = Split(pPath, "\")
    For i = 0 To UBound(x)
        strPath = strPath & x(i) & "\"
        If Not Folder_Exists(strPath) Then MkDir strPath
    Next i
End Sub

' Replaces invalid characters in a filename with a safe replacement character
Function File_SanitizeName(ByVal pFilename As String, Optional pReplacement As String = "-") As String
    Dim sTemp As String

    On Error GoTo ErrHandler

    sTemp = pFilename

    ' Replace characters that are not allowed in Windows filenames
    sTemp = Replace(sTemp, "\", pReplacement)
    sTemp = Replace(sTemp, "/", pReplacement)
    sTemp = Replace(sTemp, ":", pReplacement)
    sTemp = Replace(sTemp, """", pReplacement)
    sTemp = Replace(sTemp, "*", pReplacement)
    sTemp = Replace(sTemp, "?", pReplacement)
    sTemp = Replace(sTemp, "<", pReplacement)
    sTemp = Replace(sTemp, ">", pReplacement)
    sTemp = Replace(sTemp, "|", pReplacement)

    File_SanitizeName = sTemp
    Exit Function

ErrHandler:
    ' If something goes wrong, return a generic safe filename
    File_SanitizeName = "invalid_filename"
End Function

' Returns a collection of file paths in the specified folder that match the given file mask
Public Function File_List(ByVal pFolder As String, ByVal pMask As String) As Collection
    Dim folderPath As String
    Dim filename As String
    Dim txtFilesCollection As New Collection

    On Error GoTo ErrHandler

    ' Ensure folder path ends with a backslash
    folderPath = pFolder
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Get the first matching file
    filename = Dir(folderPath & pMask)

    ' Loop through all matching files
    Do While filename <> ""
        txtFilesCollection.Add folderPath & filename
        filename = Dir
    Loop

    ' Return the collection
    Set File_List = txtFilesCollection
    Exit Function

ErrHandler:
    Set File_List = Nothing
End Function

' Wrapper...  I keep forgetting the new name....
Function Folder_Exists(ByVal pFolder As String)
    Folder_Exists = IsFolder(pFolder)
End Function

' Wrapper...  I keep forgetting the new name....
Function File_Exists(ByVal pFile As String)
    File_Exists = IsFile(pFile)
End Function

