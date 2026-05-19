Attribute VB_Name = "libWorksheets"
' 18 May 2026
'   Rationalised namespace and moved routines from libFiles

Option Explicit
Option Private Module

' For use when bulk converting PDFs
Public Sub Worksheet_ExportCurrentAsPDF()
    Dim sFilename As String
    
    sFilename = Text_Replace(Workbook_ActiveLocalFilename, ".xlsx", ".pdf")
    sFilename = Text_Replace(sFilename, ".xls", ".pdf")
    sFilename = Text_Replace(sFilename, ".pdf", ActiveSheet.Name & ".pdf")
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=sFilename, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

' Exports the current worksheet to a new XLSX workbook
Public Sub Worksheet_ExportCurrentAsXLSX(Optional AFilename As String = "")
    Dim sFilename As String, sFolder As String
    Dim oNew As Workbook, oCurrent As Workbook, oSheet As Worksheet
    Dim i As Long

    On Error GoTo ErrHandler

    sFolder = Path_AddTrailingDelimiter(Path_GetFolder(AFilename))
    If sFolder = "\" Then sFolder = Workbook_ActiveDirectory

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
Public Sub Worksheet_ExportCurrentAsCSV(Optional AIncludeSheetName As Boolean = True, Optional AFilename As String = "")
    Dim sFilename As String, sFilenameOnly As String
    Dim sCSV As String, sExt As String, sSheetname As String

    On Error GoTo ErrHandler

    sFilename = Workbook_ActiveLocalFilename
    sFilenameOnly = Path_GetFileNameNoExt(sFilename)
    sExt = LCase(Path_GetExtension(sFilename))
    sSheetname = ActiveSheet.Name

    If AFilename <> "" Then
        sCSV = Workbook_ActivePath & AFilename
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

Public Sub Worksheet_FreezeTopRow(ByVal pWS As Worksheet)
    pWS.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

Public Function Worksheet_Add(ByVal pWB As Workbook, ByVal pName As String, Optional ByVal pIndex As Long = 1) As Worksheet
    Dim ws As Worksheet
    
    Set ws = Worksheet_Find(pWB, pName)
    
    If ws Is Nothing Then
        Set ws = pWB.Worksheets.Add(After:=pWB.Worksheets(pWB.Worksheets.Count))
        ws.Name = pName
    End If
    
    Set Worksheet_Add = ws
    
    If pWB.Worksheets.Count = 1 Then Exit Function
    
    If pIndex > 1 Then
        If pIndex <= pWB.Worksheets.Count Then
            Worksheet_Add.Move After:=pWB.Worksheets(pIndex)
        Else
            Worksheet_Add.Move After:=pWB.Worksheets(pWB.Worksheets.Count)
        End If
    Else
        Worksheet_Add.Move Before:=pWB.Worksheets(1)
    End If
End Function

Public Sub Worksheet_Delete(ByVal pWB As Workbook, ByVal pWS As Worksheet)
    Dim oldDisplayAlerts As Boolean
    
    If pWB Is Nothing Then Exit Sub
    If pWS Is Nothing Then Exit Sub
    If pWS.Parent Is Nothing Then Exit Sub
    If Not pWS.Parent Is pWB Then Exit Sub
    If pWB.Worksheets.Count <= 1 Then Exit Sub
    
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    On Error GoTo CleanUp
    pWS.Delete
    
CleanUp:
    Application.DisplayAlerts = oldDisplayAlerts
End Sub

Public Sub Worksheet_SortTabs(ByVal pWB As Workbook)
    Dim i As Long
    Dim bSorted As Boolean
    
    If pWB Is Nothing Then Exit Sub
    If pWB.Sheets.Count < 2 Then Exit Sub
    
    ' Bubble sort
    Do
        bSorted = True
        
        For i = pWB.Sheets.Count - 1 To 1 Step -1
            If StrComp(pWB.Sheets(i).Name, pWB.Sheets(i + 1).Name, vbTextCompare) > 0 Then
                pWB.Sheets(i).Move After:=pWB.Sheets(i + 1)
                bSorted = False
            End If
        Next i
    Loop Until bSorted
End Sub

Public Function Worksheet_Exists(ByVal pWB As Workbook, ByVal pSheetName As String) As Boolean
    Dim ws As Worksheet

    Worksheet_Exists = False

    For Each ws In pWB.Worksheets
        If StrComp(ws.Name, pSheetName, vbTextCompare) = 0 Then
            Worksheet_Exists = True
            Exit Function
        End If
    Next ws
End Function

Public Function Worksheet_Find(ByVal pWB As Workbook, ByVal pName As String) As Worksheet
    On Error Resume Next
    Set Worksheet_Find = pWB.Worksheets(pName)
    On Error GoTo 0
End Function

Public Function Worksheet_IsNameValid(ByVal pName As String) As Boolean
    Dim vBadChars
    Dim i As Long

    Worksheet_IsNameValid = False

    If Len(Trim$(pName)) = 0 Then Exit Function
    If Len(pName) > 31 Then Exit Function

    vBadChars = Array(":", "\", "/", "?", "*", "[", "]")

    For i = LBound(vBadChars) To UBound(vBadChars)
        If InStr(pName, vBadChars(i)) > 0 Then Exit Function
    Next i

    If Left$(pName, 1) = "'" Then Exit Function
    If Right$(pName, 1) = "'" Then Exit Function

    Worksheet_IsNameValid = True
End Function


