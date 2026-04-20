Attribute VB_Name = "LibraryWorkSheets"
Public Function Add_Sheet(sName As String, Optional iIndex As Integer = 1) As Worksheet
    If Not Sheet_Exists(sName) Then
        Sheets.Add.Name = sName
    End If
    
    Set Add_Sheet = Sheets(sName)
    
    If iIndex > 1 Then
        If iIndex <= Sheets.Count Then
            Add_Sheet.Move After:=Sheets(iIndex)
        Else
            Add_Sheet.Move After:=Sheets(Sheets.Count)
        End If
    Else
        Add_Sheet.Move Before:=Sheets(1)
    End If
End Function

Public Sub Delete_Sheet(ASheet As Worksheet)
    On Error Resume Next
    If (Not ASheet Is Nothing) And (Sheets.Count > 1) Then
        Application.DisplayAlerts = False
        ASheet.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Public Sub Sort_Sheets()
    '
    ' This is written as a helper function for Tidy_Event_Export
    '
    
    Dim i As Integer
    Dim bSorted As Boolean
    
    bSorted = False
    
    While Not bSorted
        bSorted = True
        For i = Sheets.Count - 1 To 1 Step -1
            If Sheets(i).Name > Sheets(i + 1).Name Then
                Sheets(i).Move After:=Sheets(i + 1)
                bSorted = False
            End If
        Next i
    Wend
End Sub

Public Function Sheet_Exists(sName As String) As Boolean
    '
    ' Here as I can't find the official check
    '
    Dim i As Integer
    Dim bFound As Boolean
    
    bFound = False
    i = 1
    
    While (i <= Sheets.Count) And Not bFound
        bFound = UCase(sName) = UCase(Sheets(i).Name)
        i = i + 1
    Wend
    
    Sheet_Exists = bFound
End Function

Public Sub Swap_Sheets(iSheet1 As Integer, iSheet2 As Integer)
    Dim sName1, sName2 As String
    
    sName1 = Sheets(iSheet1).Name
    sName2 = Sheets(iSheet2).Name
    
    Sheets(sName2).Move Before:=Sheets(sName1)
    Sheets(sName1).Move After:=Sheets(iSheet2)
End Sub

Public Function FindSheet(sName As String) As Worksheet
    On Error Resume Next
    
    Set FindSheet = ActiveWorkbook.Sheets(sName)
    
    'If FindSheet Is Nothing Then
    '    For Each Sheet In Worksheets
    '        If UCase(sName) = UCase(Sheet.Name) Then
    '            Set FindSheet = Sheet
    '            Exit Function
    '        End If
    '    Next Sheet
    'End If
End Function

