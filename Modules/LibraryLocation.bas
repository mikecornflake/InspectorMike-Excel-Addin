Attribute VB_Name = "LibraryLocation"
' Remnant Code - HuizHou 2004

Option Explicit

Dim iLoc As Integer

Private Sub SetLocationColumn(iCol As Integer)
    iLoc = iCol
End Sub

Private Sub ProcessTreeNameColumn()
    Dim i As Integer
    Dim sTreename As String
    Dim iCount As Integer
    Dim sComponent_ID As String
    
    Columns(4).Select
    
    Selection.Insert Shift:=xlToRight
    
    Cells(1, 4).Value = "Component_ID"
    
    FindExtents
    ConnectToDB
        
    On Error Resume Next
    
    For i = 2 To FLastRow
        sTreename = Cells(i, 3).Value
            
        iCount = 0
        
        RunQuery ("Select * From Component Where Treename = '" + Trim(sTreename) + "'")
        
        db_Records.MoveFirst
        While Not db_Records.EOF
            iCount = iCount + 1
            sComponent_ID = db_Records.fields("Component_ID").Value
            db_Records.MoveNext
        Wend
        
        If iCount > 0 Then
            Cells(i, 4).Value = sComponent_ID
        Else
            Cells(i, 4).Value = "Doesn't Exist"
        End If
    Next i
        
    CloseConnection
End Sub

Private Sub ProcessLocationColumn()
    Dim i As Integer
    Dim sOriginal As String
    Dim sStructure As String
    
    ' Here to simplify debugging
    If iLoc = 0 Then
        iLoc = 2
        MsgBox "Location column defaulted to " & iLoc & vbCrLf & "Call SetLocationColumn(-1) to avoid this"
    End If
    
    If iLoc <> -1 Then
        InsertColumns
        FindExtents
        ConnectToDB
        
        For i = 2 To FLastRow
            sStructure = Cells(i, 1).Value
            sOriginal = UCase(Cells(i, iLoc).Value)
            
            Cells(i, iLoc + 1).Value = LookupTreenameFromLocation(sStructure, sOriginal)
            If Cells(i, iLoc + 1).Value = "" Then
                Cells(i, iLoc + 1).Value = LookupViaHeuristic(sStructure, sOriginal)
            End If
        Next i
        
        CloseConnection
    End If
End Sub

Private Function FindComponentID(sTreename As String) As Integer
  RunQuery ("Select * From Component Where Treename = '" + Trim(sTreename) + "'")
  
  db_Records.MoveFirst
  If Not db_Records.EOF Then
    FindComponentID = db_Records.fields("Component_ID").Value
  Else
    FindComponentID = -1
  End If
End Function

Private Function LookupViaHeuristic(sStructure As String, ByVal sOriginal As String) As String
    Dim sTemp As String
    Dim sTemp2 As String
    Dim i, j As Integer
    Dim oParts
    Dim sPart As String
    
    Dim iOrientation As Integer
    Dim iNumber As Integer
    Dim sCompType As String
    
    iOrientation = -1
    iNumber = -1
    sCompType = ""
    sTemp = ""
            
    sOriginal = " " + SwapString(sOriginal, ".", " ") ' :-)  For search at start of string...
    
    ' eliminate any annoying double spaces
    While InStr(sOriginal, "  ") > 0
        sOriginal = SwapString(sOriginal, "  ", " ")
    Wend
    
    sOriginal = SwapString(sOriginal, " N1", " NODE 1")
    sOriginal = SwapString(sOriginal, " N2", " NODE 2")
    sOriginal = SwapString(sOriginal, " N3", " NODE 3")
    sOriginal = SwapString(sOriginal, " N4", " NODE 4")
    sOriginal = SwapString(sOriginal, " N5", " NODE 5")
    sOriginal = SwapString(sOriginal, " N6", " NODE 6")
    sOriginal = SwapString(sOriginal, " N7", " NODE 7")
    sOriginal = SwapString(sOriginal, " N8", " NODE 8")
    sOriginal = SwapString(sOriginal, " N9", " NODE 9")
    
    sOriginal = SwapString(sOriginal, "A1-", "A1 - ")
    sOriginal = SwapString(sOriginal, "A2-", "A2 - ")
    sOriginal = SwapString(sOriginal, "A3-", "A3 - ")
    sOriginal = SwapString(sOriginal, "A4-", "A4 - ")
    sOriginal = SwapString(sOriginal, "B1-", "B1 - ")
    sOriginal = SwapString(sOriginal, "B2-", "B2 - ")
    sOriginal = SwapString(sOriginal, "B3-", "B3 - ")
    sOriginal = SwapString(sOriginal, "B4-", "B4 - ")
    
    ' the above may have added an unwanted NODE
    sOriginal = SwapString(sOriginal, "NODE NODE", "NODE")
    
    sOriginal = SwapString(sOriginal, "(", " ")
    sOriginal = SwapString(sOriginal, ")", " ")
            
    sOriginal = SwapString(sOriginal, " R4_", " R-4_")
            
    ' First we guarantee that there are spaces between the alpha's and the number's
    sOriginal = SwapString(sOriginal, "HDM", " HDM ")
    sOriginal = SwapString(sOriginal, "HOM", " HOM ")
    sOriginal = SwapString(sOriginal, "PILE", " VOM ")
    sOriginal = SwapString(sOriginal, "VDM", " VDM ")
    sOriginal = SwapString(sOriginal, "VOM", " VOM ")
    sOriginal = SwapString(sOriginal, "MEMBER", "MEMBER ")
    
    'ummm
    sOriginal = SwapString(sOriginal, "ANODE", "ZZZZZ")
    sOriginal = SwapString(sOriginal, "NODE", " NODE ")
    sOriginal = SwapString(sOriginal, "ZZZZZ", "ANODE")
    
    sOriginal = SwapString(sOriginal, ":", " ")
    sOriginal = SwapString(sOriginal, " - M1", " - MEMBER 1")
    sOriginal = SwapString(sOriginal, " - M4", " - MEMBER 4")
    sOriginal = SwapString(sOriginal, "LEG", " LEG ")
    sOriginal = SwapString(sOriginal, "SECTION", " SECTION ")
    sOriginal = SwapString(sOriginal, "CLAMP", " CLAMP ")
    sOriginal = SwapString(sOriginal, "CONDUCTOR", " CONDUCTOR ")
    sOriginal = SwapString(sOriginal, "CONDT ", " CONDUCTOR ")
    sOriginal = SwapString(sOriginal, "SECTION S", "SECTIONS")
            
    ' now we apply some simple rules
    sOriginal = SwapString(sOriginal, "MEMBER", "HOM")
    sOriginal = SwapString(sOriginal, "  _", "_")
    sOriginal = SwapString(sOriginal, " _", "_")
    sOriginal = SwapString(sOriginal, "CONDUCTOR  GUIDE FRAME", "CGF")
    sOriginal = SwapString(sOriginal, "VM", "VOM")
    sOriginal = SwapString(sOriginal, " NO. ", " ")
    sOriginal = SwapString(sOriginal, " NO ", " ")
    sOriginal = SwapString(sOriginal, " NUM ", " ")
    sOriginal = SwapString(sOriginal, " #", " ")
    sOriginal = SwapString(sOriginal, " 'S ", " ")
    'sOriginal = SwapString(sOriginal, " -", " ")
            
    ' Now, this has probably put some double spaces in the string
    ' if so, lets get rid of them
    While InStr(sOriginal, "  ") > 0
        sOriginal = SwapString(sOriginal, "  ", " ")
    Wend
            
    oParts = Split(sOriginal, " ")
            
    For j = 0 To UBound(oParts)
        sPart = Trim(oParts(j))
                
        If sPart = "HDM" Or sPart = "HOM" Or sPart = "VDM" Or sPart = "HM" Then
            If (j <= UBound(oParts)) And (IsNumber(Trim(oParts(j + 1)))) Then
                iOrientation = j
                iNumber = j + 1
                sCompType = "Member"
                        
                sTemp = oParts(iNumber)
                    
                If InStr(sTemp, "-") Then
                    ' Discard all the text following the - (including the -)
                    oParts(iNumber) = Left(sTemp, InStr(sTemp, "-") - 1)
                End If
                sTemp = "" ' Need to set this to blank cause I've just used a variable from
                                   ' somewhere else rather than declare a new variable
            End If
        ElseIf sPart = "VOM" Then
            If (j <= UBound(oParts)) And (IsNumber(Trim(oParts(j + 1)))) Then
                iOrientation = j
                iNumber = j + 1
                If InStr(sOriginal, "LEG") Then  ' They call leg Sections VOM, we call them Leg Sections
                    sCompType = "Leg Section"
                Else
                    sCompType = "Member"
                End If
            End If
        ElseIf sPart = "NODE" Then
            If (j <= UBound(oParts)) And (IsNumber(Trim(oParts(j + 1)))) Then
                iOrientation = j
                iNumber = j + 1
                sCompType = "Node"
                    
                Exit For
            End If
        ElseIf sPart = "LEG" Then
            iOrientation = j
            iNumber = j + 1
                    
            sCompType = "Leg"
        ElseIf sPart = "SECTION" Then
            iOrientation = j
            iNumber = j + 1
                    
            sCompType = "Leg Section"
        ElseIf sPart = "CLAMP" Then
            iOrientation = j
            iNumber = j + 1
            sCompType = "Clamp"
        ElseIf sPart = "CONDUCTOR" Then
            iOrientation = j
            iNumber = j + 1
                    
            sCompType = "Conductor"
        End If
    Next j
                
    If (iOrientation <> -1) And (UBound(oParts) > 0) Then
        sTemp2 = oParts(iNumber)
        sTemp2 = RemoveSubString(sTemp2, ".")
        sTemp2 = SwapString(sTemp2, "_EL_", " EL ") ' Crafty sod, _ force grouping of names in Excel files
        oParts(iNumber) = sTemp2
                
        sTemp = oParts(iOrientation) & " " & oParts(iNumber)
                    
        If sCompType = "Leg" Then
            sTemp2 = oParts(iNumber)
            oParts(iNumber) = RemoveSubString(sTemp2, "-")
        End If
                            
        LookupViaHeuristic = LookupTreename("Huizhou / " + sStructure, " " & oParts(iNumber), sCompType)
        If LookupViaHeuristic = "" Then
            LookupViaHeuristic = LookupTreename("Huizhou / " + sStructure, "-" & oParts(iNumber), sCompType)
        End If
    Else
        LookupViaHeuristic = ""
    End If
End Function

Private Sub InsertColumns()
    If iLoc <> -1 Then
        Columns(iLoc + 1).Select
    
        Selection.Insert Shift:=xlToRight
    
        Cells(1, iLoc).Value = "Original " & Cells(1, iLoc).Value
        Cells(1, iLoc + 1).Value = "Treename"
    End If
End Sub
