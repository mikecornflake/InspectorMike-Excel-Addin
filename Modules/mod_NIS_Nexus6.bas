Attribute VB_Name = "mod_NIS_Nexus6"
Private FPipeline As Boolean
Private FMaxFindingCols As Long
Dim FMultimedia As Worksheet
Dim FmmFilenameCol As Long, FmmNewFilenameCol As Long, FmmFolderCol As Long

Sub Test()
    Dim oWorkbook As Workbook
    
    Set oWorkbook = ActiveWorkbook
        
    'oWorkbook.Activate
    'Duplicate_ActiveBook ("D:\Temp\Working\2956-BKA200-MKA (Oil).xlsx")
    'ProcessNexus6EventExport ("Gippsland Basin / Pipelines / BKA200-MKA (Oil)")
    'ActiveWorkbook.Save
    'ActiveWorkbook.Close
    
    oWorkbook.Activate
    Duplicate_ActiveBook ("D:\Temp\Working\5355-CBA150-HLA (Oil).xlsx")
    ProcessNexus6EventExport ("Gippsland Basin / Pipelines / CBA150-HLA (Oil)")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    oWorkbook.Activate
    Duplicate_ActiveBook ("D:\Temp\Working\5358-01 - Oil To KFB.xlsx")
    ProcessNexus6EventExport ("Gippsland Basin / Platforms / Kingfish A (KFA) / Risers / 01 - Oil To KFB")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    oWorkbook.Activate
    Duplicate_ActiveBook ("D:\Temp\Working\5359-06 - Oil From WKF.xlsx")
    ProcessNexus6EventExport ("Gippsland Basin / Platforms / Kingfish A (KFA) / Risers / 06 - Oil From WKF")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    oWorkbook.Activate
    Duplicate_ActiveBook ("D:\Temp\Working\5360-14 - Fuel Gas From MLA.xlsx")
    ProcessNexus6EventExport ("Gippsland Basin / Platforms / Kingfish A (KFA) / Risers / Fuel Gas Caisson / 14 - Fuel Gas From MLA")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    oWorkbook.Activate
    Duplicate_ActiveBook ("D:\Temp\Working\5362-01 - Oil To KFA.xlsx")
    ProcessNexus6EventExport ("Gippsland Basin / Platforms / West Kingfish (WKF) / Risers / 01 - Oil To KFA")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    oWorkbook.Activate
    Duplicate_ActiveBook ("D:\Temp\Working\5356-Seahorse (SHA).xlsx")
    ProcessNexus6EventExport ("Gippsland Basin / Subsea Completions / Seahorse (SHA)")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    oWorkbook.Activate
    Duplicate_ActiveBook ("D:\Temp\Working\5357-Tarwhine (TWA).xlsx")
    ProcessNexus6EventExport ("Gippsland Basin / Subsea Completions / Tarwhine (TWA)")
    ActiveWorkbook.Save
    ActiveWorkbook.Close
   
    MsgBox "Finished"
End Sub

Sub Test2()
    ' Only works on the avtive Worksheet
    'Delete_Events_By_Location ("Gippsland Basin / Pipelines / BMA350-VS3 (Gas)")
    FormatColumns
End Sub

Sub Test3()
    Duplicate_ActiveBook ("D:\Temp\Working\Delete Me.xlsx")
    ProcessWorkbook ("")
End Sub

Sub ProcessNexus6EventExport(ALocation As String)
    Dim oHUVR As Worksheet, oWorkbook As Workbook
    
    Dim iFilterCol As Long, iFilenameCol As Long, iFolderCol As Long
    Dim sFilter As String, sFilename As String, sFolder As String
    Dim iRow As Long

    Set oWorkbook = ActiveWorkbook
    
    Set oHUVR = FindSheet("HUVR")
    If oHUVR Is Nothing Then
        DoProcessNexus6EventExport (ALocation)
    Else
        oHUVR.Activate
        
        If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
        
        ForceFindExtents
        
        iFilterCol = FindColumn("Asset Location.Full Location")
        iFilenameCol = FindColumn("Destination Filename")
        iFolderCol = FindColumn("Destination Folder")
        
        For iRow = 2 To FLastRow
            oWorkbook.Activate
            oHUVR.Cells(iRow, iFilterCol).Select
            
            sFilter = oHUVR.Cells(iRow, iFilterCol).Value
            sFilename = oHUVR.Cells(iRow, iFilenameCol).Value
            sFolder = oHUVR.Cells(iRow, iFolderCol).Value
            
            If (UCase(sFilename) <> "SKIP") And (UCase(sFolder) <> "SKIP") Then
                Duplicate_ActiveBook (AddTrailingDelimiter(sFolder) + sFilename)
                DoProcessNexus6EventExport (Trim(sFilter))
                
                ActiveWorkbook.Save
                ActiveWorkbook.Close
            End If
        Next iRow
        
    End If
End Sub

Private Sub DoProcessNexus6EventExport(ALocation As String)
    Dim oSheet As Worksheet
    Dim bStruct As Boolean
    Dim sServer As String, sDatabase As String
    Dim sUser As String, sPassword As String
    
    sServer = RegistryRead("DOF_Addin", "Nexus 6", "Server", "INS-SQL-NIC01\SQLEXPRESS")
    sDatabase = RegistryRead("DOF_Addin", "Nexus 6", "Database", "Esso_Master")
    sUser = RegistryRead("DOF_Addin", "Nexus 6", "User", "")
    sPassword = RegistryRead("DOF_Addin", "Nexus 6", "Password", "")
    FPipeline = StringBool(RegistryRead("DOF_Addin", "Nexus 6", "Pipeline", "True"))
    
    If sServer <> "" Then
        Call ConnectToSQLOLDDB(sServer, sDatabase, sUser, sPassword)
    End If
    
    Tidy_Tabs
    Find_Multimedia
    
    'Delete unwanted sheets
    For Each oSheet In ActiveWorkbook.Sheets
        If Not oSheet.Visible Then
            Call Delete_Sheet(oSheet)
        End If
    Next oSheet
    
    Set oSheet = FindSheet("Legend")
    If Not oSheet Is Nothing Then
        Call Delete_Sheet(oSheet)
    End If
    
    Set oSheet = FindSheet("HUVR")
    If Not oSheet Is Nothing Then
        Call Delete_Sheet(oSheet)
    End If
    
    For Each oSheet In ActiveWorkbook.Sheets
        If oSheet.Visible Then
            oSheet.Select
            DoEvents
            
            If (oSheet.Name <> "Findings") And (oSheet.Name <> "Multimedia") Then
                If ALocation <> "" Then
                    Delete_Events_By_Location (ALocation)
                End If
                
                If Cells(2, 1).Value = "" Then
                    Call Delete_Sheet(oSheet)
                Else
                    TidyEventSheet (bPipeline)
                End If
            End If
        End If
    Next oSheet
    
    Sort_Sheets
    
    Populate_Findings_Tab
    
    If Not FMultimedia Is Nothing Then
        Call Delete_Sheet(FMultimedia)
        Set FMultimedia = Nothing
    End If
    
    CloseConnection
    
    ActiveWorkbook.Save
End Sub

Private Sub TidyEventSheet(APipeline As Boolean)
    Dim sStatus As String
    
    ' First pass tidy - remove conditional formatting, move columns, delete unwanted
    sStatus = ActiveSheet.Name + ". Tidying sheet"
    
    Application.StatusBar = sStatus
    Remove_All_Conditional_Formatting
    
    Cells.Select
    Selection.Validation.Delete
    
    Cells(2, 1).Select
    
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    Application.StatusBar = sStatus + ".  Moving columns"
    Call Move_Column2("Multimedia.Name", "Start - Survey - Standard.Easting")
    Call Move_Column2("Multimedia.Image", "Start - Survey - Standard.Easting")
    Call Move_Column2("Workpack.Name", "Start - Survey - Standard.Easting")
    
    Application.StatusBar = sStatus + ".  Moving image links to columns"
    Normalise_Event_By_MM
    Normalise_Event_By_Colour
    Reprocess_Multimedia_By_MM_Tab
    
    Application.StatusBar = sStatus + ".  Deleting extra columns"
    Delete_Column ("Finding.Anomaly")
    Delete_Column ("Finding.Remedial Action")
    Delete_Column ("Finding.Anomaly Required")
    Delete_Column ("Finding.Severity")
    
    Delete_Column ("Event Review.Personnel")
    Delete_Column ("Event Review.Date / Time")
    Delete_Column ("Event Review.Description")
    
    Delete_Column ("Survey Set.Name")
    Delete_Column ("Survey Set.Comments")
    
    If Not APipeline Then
        Delete_Column ("Start - Survey - Pipeline.KP")
        Delete_Column ("End - Survey - Pipeline.KP")
        
        Delete_Column ("Start - Survey - Pipeline.DCC")
        Delete_Column ("End - Survey - Pipeline.DCC")
    End If
    
    Application.StatusBar = sStatus + ".  Merging Event Type and Event Number"
    AddFindingID
    MergeEventAndEventNumber
    
    BasicTidy
    FormatColumns
    
    Cells(2, 1).Select
    Application.StatusBar = ""
End Sub

Private Sub AddFindingID()
    ' This is run BEFORE MergeEventAndEventNumber, so the two columns are available
    Dim iEventNumCol As Long, iEventTypeCol As Long, iWorkpackCol As Long, iFindingCol As Long
    Dim sEventType As String, sEventNum As String, sCode As String, sWorkpack As String
    Dim sFindingID As String

    If Not ConnectedToDb Then
        Exit Sub
    End If
    
    ForceFindExtents

    iEventNumCol = Find_Column("Event.Event Number")
    iEventTypeCol = Find_Column("Event.Event Type")
    iWorkpackCol = Find_Column("Workpack.Name")
    iFindingCol = Find_Column("Finding.Code")

    If (iEventNumCol > 0) And (iEventTypeCol > 0) And (iWorkpackCol > 0) And (iFindingCol > 0) Then
        For iRow = 2 To FLastRow
            sCode = Trim(Cells(iRow, iFindingCol).Value)
            
            If sCode <> "" Then
                sEventType = Trim(Cells(iRow, iEventTypeCol).Value)
                sEventNum = Trim(Cells(iRow, iEventNumCol).Value)
                sWorkpack = Trim(Cells(iRow, iWorkpackCol).Value)
                
                sFindingID = FindingID(sEventType, sEventNum, sWorkpack, sCode)
                If sFindingID <> "" Then
                    Cells(iRow, iFindingCol).Value = sCode & "-" & sFindingID
                End If
            End If
        Next iRow
    End If
End Sub

Private Function FindingID(AEventType As String, AEventNum As String, AWorkpack As String, ACode As String) As String
    FindingID = ""
    Dim vFindingID As Variant
    Dim sTemp As String, sSQL As String
    
    
    If Not ConnectedToDb Then
        Exit Function
    End If
    
    sSQL = "Select F.Finding_ID As 'FindingID' " _
         & "From Finding F " _
         & "  Inner Join Header H on (H.Header_ID=F.Header_ID And H.Event_No='" & AEventNum & "') " _
         & "  Inner Join Table_Def TD ON (H.TD_ID=TD.TD_ID And TD.Name='" & AEventType & "') " _
         & "  Inner Join Workpack W On (W.Workpack_ID=H.Workpack_ID And W.Name='" & AWorkpack & "') " _
         & "  Inner Join Code Code On (Code.Code_ID=F.Code_ID And Code.Code='" & ACode & "') "
    
    vFindingID = QuickValue(sSQL, "FindingID")
    
    If IsNull(vFindingID) Then
        FindingID = ""
    Else
        FindingID = vFindingID
    End If
End Function


Private Sub MergeEventAndEventNumber()
    Dim iEventNumCol As Long, iEventTypeCol As Long, iEventCol As Long
    Dim sTemp As String
    Dim iRow As Long
    
    ForceFindExtents
    
    Add_Column ("Event")
    Call Move_Column2("Event", "Event.Event Number")
    iEventNumCol = Find_Column("Event.Event Number")
    iEventTypeCol = Find_Column("Event.Event Type")
    iEventCol = Find_Column("Event")
    
    If (iEventNumCol > 0) And (iEventTypeCol > 0) Then
        For iRow = 2 To FLastRow
            sTemp = Trim(Cells(iRow, iEventTypeCol).Value) & " " & Trim(Cells(iRow, iEventNumCol).Value)
            Cells(iRow, iEventCol).Value = sTemp
        Next iRow
    End If
    
    Delete_Column ("Event.Event Number")
    Delete_Column ("Event.Event Type")
End Sub

Private Sub Tidy_Tabs()
    ' Set default code on tabs and simplify the Tab Captions
    
    Dim oSheet As Worksheet
    
    For Each oSheet In ActiveWorkbook.Sheets
        If oSheet.Visible Then
            oSheet.Tab.ColorIndex = 35
            
            oSheet.Name = Trim(Replace(oSheet.Name, "(Events)", ""))
        End If
    Next oSheet
    
    Set oSheet = FindSheet("Findings")
    
    If Not oSheet Is Nothing Then
        oSheet.Tab.ColorIndex = 40
    End If
End Sub

Private Sub Copy_Finding(ADest As Worksheet, ADestRow As Long, ASource As Worksheet, ASourceRow As Long)
    Dim iDestCol As Long
    Dim iSourceCol As Long
    Dim sColumn As String
    Dim iMM As Long
    
    ASource.Select
    
    ' Process the known columns
    For iDestCol = 1 To 9
        sColumn = ADest.Cells(1, iDestCol).Value
        
        iSourceCol = Find_Column(sColumn)
        If iSourceCol <> -1 Then
            ADest.Cells(ADestRow, iDestCol).Value = ASource.Cells(ASourceRow, iSourceCol).Value
        End If
    Next iDestCol
    
    ' Process the optional Multimedia Columns
    iMM = 1
    iSourceCol = Find_Column("Multimedia " & iMM)
    
    While iSourceCol <> -1
        ADest.Cells(1, FMaxFindingCols + iMM) = "Multimedia " & iMM
        ADest.Cells(ADestRow, FMaxFindingCols + iMM).Formula = ASource.Cells(ASourceRow, iSourceCol).Formula
        
        iMM = iMM + 1
        iSourceCol = Find_Column("Multimedia " & iMM)
    Wend
End Sub

Private Function Create_Findings_Tab() As Worksheet
    Dim oFindings As Worksheet, oSheet As Worksheet

    Set oFindings = FindSheet("Findings")
    If oFindings Is Nothing Then
        Set oFindings = Add_Sheet("Findings", 1)
        
        oFindings.Select
        
        ' If more baseline columns get added or removed then don't forget to change
        ' FMaxFindingCols below...
        Add_Column ("Asset Location.Full Location")
        Add_Column ("Event")
        Add_Column ("Event.Start Clock")
        If FPipeline Then
            Add_Column ("Start - Survey - Pipeline.KP")
        End If
        Add_Column ("Start - Survey - Standard.Depth")
        Add_Column ("Finding.Code")
        Add_Column ("Finding.Reason")
        Add_Column ("Commentary.Notes")
        
        If FPipeline Then
            FMaxFindingCols = 8
        Else
            FMaxFindingCols = 7
        End If
        
        oFindings.Tab.ColorIndex = 40
        Cells(2, 1).Select
    End If
    
    Set Create_Findings_Tab = oFindings
End Function

Private Sub Populate_Findings_Tab()
    Dim oFindings As Worksheet, oSheet As Worksheet
    Dim iFindingsRow As Long, iRow As Long
    Dim iEventTypeCol As Long
    Dim iFindingEventTypeCol As Long
    
    Dim sStatus As String
    Dim iFindingCodeCol As Long, iFindingReasonCol As Long
    
    sStatus = ActiveSheet.Name + ". Populating Findings Tab: "
    
    Set oFindings = FindSheet("Findings")
    If oFindings Is Nothing Then
        Set oFindings = Create_Findings_Tab
    End If
    
    oFindings.Select
    
    ForceFindExtents
    iFindingEventTypeCol = Find_Column("Event")
    
    iFindingsRow = 2
    
    For Each oSheet In ActiveWorkbook.Sheets
        If (oSheet.Name <> "Findings") And (oSheet.Name <> "Multimedia") And (oSheet.Visible) Then
            oSheet.Select
            
            If Cells(2, 1).Value <> "" Then
                ForceFindExtents
                iFindingCodeCol = Find_Column("Finding.Code")
                iFindingReasonCol = Find_Column("Finding.Reason")
                iEventTypeCol = Find_Column("Event")

                For iRow = 2 To FLastRow
                    If iRow Mod 10 = 0 Then
                        Application.StatusBar = sStatus + "Row " & iRow & " of " & FLastRow
                    End If
                    
                    ' If this row has findings, then copy it to the Findings Sheet
                    If (Cells(iRow, iFindingCodeCol).Value <> "") Or (Cells(iRow, iFindingReasonCol).Value <> "") Then
                        Application.StatusBar = sStatus + "Row " & iRow & " of " & FLastRow & ".  Copying Finding."
                        Call Copy_Finding(oFindings, iFindingsRow, oSheet, iRow)
                        
                        ' On the Event Sheet Put in a hyperlink to the Findings Page
                        oSheet.Select
                        oSheet.Cells(iRow, iEventTypeCol).Select
                        oSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" + oFindings.Name + "'!A" & iFindingsRow, TextToDisplay:=Selection.Value
                        
                        ' On the Findings Sheet, put in the hyperlink to the Events Page
                        oFindings.Select
                        oFindings.Cells(iFindingsRow, iFindingEventTypeCol).Select
                        oFindings.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & oSheet.Name & "'!A" & iRow, TextToDisplay:=Selection.Value

                        iFindingsRow = iFindingsRow + 1
                        
                        oSheet.Select
                    End If
                Next iRow
                
                Highlight_Finding
                Cells(2, 1).Select
            End If
        End If
    Next oSheet
    
    Application.StatusBar = ""
    oFindings.Select
    
    BasicTidy
    FormatColumns
    
    Cells(2, 1).Select
End Sub

Private Sub Delete_Events_By_Location(AAllowedLocation As String)
    ' TO BE CALLED BEFORE Normalise_Event_By_MM
    Dim sStatus As String
    Dim sAsset As String
    Dim sAllowed As String
    Dim iAssetCol As Long
    Dim iRow As Long
    Dim iMMNameCol As Long, iImageCol As Long
    Dim sMMName As String, sMMImage As String
    Dim sImagePath As String
    Dim sFilename As String

    On Error Resume Next
    Application.ScreenUpdating = False
    
    sStatus = ActiveSheet.Name + ". Deleting unwanted assets: "
    ForceFindExtents
    
    iMMNameCol = Find_Column("Multimedia.Name")
    iImageCol = Find_Column("Multimedia.Image")
    
    sFilename = ActiveWorkbookLocalFilename
    sImagePath = AddTrailingDelimiter(ExtractFolder(sFilename)) + AddTrailingDelimiter(ExtractFilenameOnly(sFilename) + "_Images")
    
    iAssetCol = Find_Column("Asset Location.Full Location")
    sAllowed = UCase(AAllowedLocation)
    
    If iAssetCol <> -1 Then
        For iRow = FLastRow To 2 Step -1
                If iRow Mod 10 = 0 Then
                    Application.StatusBar = sStatus + "Row " & iRow & " of " & FLastRow
                End If
                
                sAsset = UCase(Trim(Cells(iRow, iAssetCol).Value))
                
                If InStr(sAsset, sAllowed) <> 1 Then
                    ' First delete the images listed in this row
                    sMMImage = Trim(Cells(iRow, iImageCol).Value)
                    
                    If sMMImage <> "" Then
                        Application.StatusBar = sStatus + "Row " & iRow & " of " & FLastRow & ". Deleting " + sMMImage
                        If Not DeleteFile(sImagePath + sMMImage) Then
                            ' Debug.Print ("Failed to delete " + sImagePath + sMMImage)
                        End If
                    End If
                    
                    ' Now, delete the row
                    Rows(iRow).Delete
                Else
                    Debug.Print ("Allow [" & ActiveSheet.Name & "] " & sAsset)
                End If
        Next iRow
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    Cells(2, 1).Select
End Sub

Private Sub Normalise_Event_By_MM()
    ' In the Nexus 6 event export, if there are multiple MM per event, then the event row is repeated n times
    ' this routine instead adds n "Multimedia x" columns, each containing a hyperlink.
    
    Dim sCurrEventNo As String, sActiveEventNo As String
    Dim iRow As Long
    Dim iActiveEventRow As Long
    Dim iImageColsAdded As Long
    Dim iEventNoCol As Long
    Dim iMMNameCol As Long, iImageCol As Long
    Dim sMMName As String, sMMImage As String
    Dim iFindingCol As Long
    Dim sFormula As String
    Dim bDebug As Boolean
    Dim bIsFinding As Boolean
    Dim sImagePath As String, sFindingImagePath As String
    Dim sFilename As String
    
    Dim sStatus As String
    
    sStatus = ActiveSheet.Name + ". Moving images from rows to columns: "
    
    bDebug = False
    
    ForceFindExtents
    
    iEventNoCol = Find_Column("Event.Event Number")
    iMMNameCol = Find_Column("Multimedia.Name")
    iImageCol = Find_Column("Multimedia.Image")
    iFindingCol = Find_Column("Finding.Code")
    
    sFilename = ActiveWorkbookLocalFilename
    
    sImagePath = AddTrailingDelimiter(ExtractFolder(sFilename)) + AddTrailingDelimiter(ExtractFilenameOnly(sFilename) + "_Images")
    
    If FMultimedia Is Nothing Then
        sFindingImagePath = AddTrailingDelimiter(ExtractFolder(sFilename)) + AddTrailingDelimiter(ExtractFilenameOnly(sFilename) + "_Finding_Images")
        ForceDirectories (sFindingImagePath)
    End If
    
    Application.StatusBar = sStatus + "Beginning first pass"
    
    ' Do the prerequisite columns exist (they won't if this is a second accidental run over same data)
    If (iMMNameCol <> -1) And (iEventNoCol <> -1) And (iImageCol <> -1) Then
        ' This code assumes the original sort order is in place
        ' Two parses over the table.
        '    First pass we add the additional image columns
        '    Second pass we delete the now duplicate rows
        
        sCurrEventNo = ""
        sActiveEventNo = ""
        iImageColsAdded = -1
        iActiveEventRow = -1
            
        ' First pass - move the data from the rows to new columns
        For iRow = 2 To FLastRow
            If iRow Mod 10 = 0 Then
                Application.StatusBar = sStatus + "First pass:  Row " & iRow & " of " & FLastRow
            End If
            If bDebug Then
                Cells(iRow, iEventNoCol).Select
            End If
            
            sCurrEventNo = Cells(iRow, iEventNoCol).Value
            If sCurrEventNo <> sActiveEventNo Then
                sActiveEventNo = sCurrEventNo
                iActiveEventRow = iRow
                
                bIsFinding = Cells(iRow, iFindingCol).Value <> ""
            End If
            
            ' Is there an image here?
            If bDebug Then
                Cells(iRow, iMMNameCol).Select
            End If
            sMMName = Cells(iRow, iMMNameCol).Value
            sMMImage = Trim(Cells(iRow, iImageCol).Value)
            
            If sMMImage <> "" Then
                ' Do we need to add an image column?
                If (iRow - iActiveEventRow) > iImageColsAdded Then
                    Call Insert_Column("Multimedia " & (iRow - iActiveEventRow) + 1, iImageCol + (iRow - iActiveEventRow) + 1, False)
                    iImageColsAdded = iImageColsAdded + 1
                    
                    Columns(iImageCol + (iRow - iActiveEventRow) + 1).Select
                    'Selection.Font.Color = -65536
                End If
                
                sFormula = Cells(iRow, iImageCol).Formula
                If sMMName <> "" Then
                    sFormula = Replace(sFormula, """" + sMMImage + """", """" + sMMName + """")
                End If
                
                Cells(iActiveEventRow, iImageCol + (iRow - iActiveEventRow) + 1).Formula = sFormula
                
                If bIsFinding And (FMultimedia Is Nothing) Then
                    If FileExists(sImagePath & sMMImage) Then
                        On Error Resume Next
                        Call FileCopy(sImagePath & sMMImage, sFindingImagePath & "\" & sMMImage)
                    End If
                End If
                
                If iRow <> iActiveEventRow Then
                    Cells(iRow, 1).Value = "Delete"
                End If
            End If
        Next iRow
        
        ' Second Pass.  Going to delete the "duplicate image rows"
        ' Nexus colour codes the duplicate rows - don't want to accidentally delete the "duplicate Finding Rows")
        For iRow = FLastRow To 2 Step -1
            If bDebug Then
                Cells(iRow, 1).Select
            End If
            
            If iRow Mod 10 = 0 Then
                Application.StatusBar = sStatus + "Second pass:  Row " & iRow & " of " & FLastRow
            End If
            
            If (Cells(iRow, 1).Value = "Delete") Then
                Rows(iRow).Delete
            End If
        Next iRow
        
        ' Delete the original image columns
        Columns(iImageCol).Delete
        Columns(iMMNameCol).Delete
    End If
    
    ' And finally - end neatly :-)
    Cells(2, 1).Select
    Application.StatusBar = ""
End Sub

Private Sub Normalise_Event_By_Colour()
    Dim bHasSubEvents As Boolean
    Dim bHasDuplicateAssets As Boolean
    Dim bHasOther As Boolean
    Dim bHasReview As Boolean
    Dim iRow As Long
    Dim iEventCol As Long
    Dim iTemp As Long
    
    ForceFindExtents
    
    bHasSubEvents = False
    bHasDuplicateAssets = False
    bHasOther = False
    bHasReview = False
    
    For iRow = 2 To FLastRow
        iTemp = Cells(iRow, 1).Interior.ColorIndex
        bHasSubEvents = bHasSubEvents Or (iTemp = 15)
        bHasDuplicateAssets = bHasDuplicateAssets Or (iTemp = 20)
        bHasReview = bHasReview Or (iTemp = 42)
        
        bHasOther = bHasOther Or ((iTemp <> 42) And (iTemp <> 15) And (iTemp <> 20) And (iTemp <> -4142))
    Next iRow
    
    If bHasSubEvents And bHasDuplicateAssets Then
        ' Too hard
        
        Exit Sub
    End If
    
    iEventCol = FindFirstColumn(Array("Event", "Event.Event Number"))
    
    If bHasSubEvents Then
        Cells(1, 1).Select
        
        For iRow = 2 To FLastRow
            If (Cells(iRow, 1).Interior.ColorIndex = 15) Then
                Rows(iRow).Select
                With Selection.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next iRow
        
        Cells(1, 1).Select
    End If
    
    If (bHasDuplicateAssets Or bHasReview) And (iEventCol > 0) And (Not bHasOther) Then
        ForceFindExtents
        Call SelectTable("A1", FLastRow, FLastColumn)
        Application.DisplayAlerts = False
        Selection.RemoveDuplicates Columns:=iEventCol, Header:=xlYes
        Application.DisplayAlerts = True
        
        For iRow = 2 To FLastRow
            If (Cells(iRow, 1).Interior.ColorIndex = 20) Then
                Rows(iRow).Select
                With Selection.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next iRow
        Cells(1, 1).Select
    End If
End Sub

Private Sub Highlight_Finding()
    ' Go through the worksheet, and highlight findings by light orange
    Dim sStatus As String
    Dim iRow As Long
    Dim iFindingCodeCol As Long, iFindingReasonCol As Long, iEventCol As Long
    Dim sLastEvent As String, sCurrEvent As String
    Dim bLastEventWasFinding As Boolean
    Dim bCurrFinding As Boolean
    
    sStatus = ActiveSheet.Name + ". Highlighting Finding Rows: "
    Application.StatusBar = sStatus + "Initialising"
    ForceFindExtents
    
    iFindingCodeCol = Find_Column("Finding.Code")
    iFindingReasonCol = Find_Column("Finding.Reason")
    iEventCol = FindFirstColumn(Array("Event", "Event.Event Number"))
    sLastEvent = ""
    bLastEventWasFinding = False
    bCurrFinding = False
    
    If (iFindingCodeCol <> -1) And (iFindingReasonCol <> -1) And (iEventCol <> 0) Then
        For iRow = 2 To FLastRow
            If iRow Mod 10 = 0 Then
                Application.StatusBar = sStatus + "Row " & iRow & " of " & FLastRow
            End If
            
            bCurrFinding = (Cells(iRow, iFindingCodeCol).Value <> "") Or (Cells(iRow, iFindingReasonCol).Value <> "")
            If (bCurrFinding) Then
                Rows(iRow).Select
                
                ' Light Orange
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
                
                ActiveSheet.Tab.ColorIndex = 40
                
                bLastEventWasFinding = True
            End If
            
            sCurrEvent = Cells(iRow, iEventCol).Value
            
            If (sCurrEvent = sLastEvent) And (bLastEventWasFinding) Then
                Rows(iRow).Select
                
                ' Light Orange
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
            End If
            
            If (sCurrEvent <> sLastEvent) Then
                sLastEvent = sCurrEvent
                bLastEventWasFinding = bCurrFinding
            End If
        Next iRow
    End If
    
    Cells(2, 1).Select
    Application.StatusBar = ""
End Sub

Private Sub Find_Multimedia()
    Set FMultimedia = FindSheet("Multimedia")
    
    If Not FMultimedia Is Nothing Then
        FMultimedia.Select
        
        FmmFilenameCol = Find_Column("Filename")
        
        FmmNewFilenameCol = FindFirstColumn(Array("New_Filename", "New Filename"))
        FmmFolderCol = FindFirstColumn(Array("Recording_Folder", "Recording Folder"))
    End If
End Sub

Private Sub Reprocess_Multimedia_By_MM_Tab()
    Dim oCurrent As Worksheet
    Dim iMediaCol As Long, iColAdd As Long
    Dim sRootFolder As String, sImagesFolder As String
    Dim sNewLink As String
    Dim sOrigFilename As String
    Dim sNewFilename As String, sNewFolder As String
    Dim immRow As Long
    Dim sName As String, sTemp As String, sExt As String
    Dim sSearchFilename As String
    Dim sDisplay As String
    
    ' Got an oddity with some files being reported as "Not Found" when they're clearly there.
    ' This will allow those files to be manually processed
    On Error Resume Next
    
    If FMultimedia Is Nothing Then
        Exit Sub
    End If
    
    Set oCurrent = ActiveWorkbook.ActiveSheet
    iMediaCol = Find_Column("Multimedia 1")
    
    If iMediaCol <= 0 Then
        Exit Sub
    End If
    
    sImagesFolder = ExtractFilenameOnly(ActiveWorkbookLocalFilename) & "_Images\"
    sRootFolder = ActiveWorkbookPath & sImagesFolder
    
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    
    For iRow = 2 To oCurrent.UsedRange.Rows.Count
        iColAdd = 0
        
        While (Left(Trim(oCurrent.Cells(1, iMediaCol + iColAdd).Value), 11) = "Multimedia ")
            sDisplay = Trim(oCurrent.Cells(iRow, iMediaCol + iColAdd).Value)
            sTemp = Trim(oCurrent.Cells(iRow, iMediaCol + iColAdd).Formula)
            
            sOrigFilename = StringBetween(sTemp, "Images\", """,")
            sName = StringBeforeLast(ExtractFilenameOnly(sOrigFilename), "_")
            sExt = ExtractFilenameExt(sOrigFilename)
            
            sSearchFilename = sName + sExt
            
            If sOrigFilename <> "" Then
                immRow = Find_In_Column(FmmFilenameCol, sSearchFilename, FMultimedia)
                
                If immRow > 0 Then
                    sNewFolder = Trim(FMultimedia.Cells(immRow, FmmFolderCol).Value)
                    
                    ' sNewFilename = sOrigFilename
                    
                    sNewFilename = Format(immRow - 1, "0000") & " - " & ValidateFilename(Trim(FMultimedia.Cells(immRow, FmmNewFilenameCol).Value))
                    
                    If sNewFolder = "DO NOT MOVE" Then
                        sNewLink = Replace(sImagesFolder & sNewFilename, " ", "%20")
                        If FileExists(sRootFolder & sOrigFilename) Then
                            ' ForceDirectories (AddTrailingDelimiter(sRootFolder & sNewFolder))
                            
                            Call FileCopy(sRootFolder & sOrigFilename, sRootFolder & sNewFilename)
                            If FileExists(sRootFolder & sNewFilename) Then
                                DeleteFile (sRootFolder & sOrigFilename)
                            End If
                        End If
                        
                        ' It's possible the image was already processed on a different tab, if so, this hyperlink still needs updating
                        If FileExists(sRootFolder & sNewFilename) Then
                            sTemp = "=Hyperlink("""", """")"
                            oCurrent.Cells(iRow, iMediaCol + iColAdd).Select
                            
                            Selection.Hyperlinks(1).Address = sNewLink
                            Selection.Hyperlinks(1).TextToDisplay = sNewFilename
                        End If
                    Else
                        sNewLink = Replace(sImagesFolder & AddTrailingDelimiter(sNewFolder) & sNewFilename, " ", "%20")
                        If FileExists(sRootFolder & sOrigFilename) Then
                            ForceDirectories (AddTrailingDelimiter(sRootFolder & sNewFolder))
                            
                            Call FileCopy(sRootFolder & sOrigFilename, sRootFolder & AddTrailingDelimiter(sNewFolder) & sNewFilename)
                            If FileExists(sRootFolder & AddTrailingDelimiter(sNewFolder) & sNewFilename) Then
                                DeleteFile (sRootFolder & sOrigFilename)
                            End If
                        End If
                        
                        ' It's possible the image was already processed on a different tab, if so, this hyperlink still needs updating
                        If FileExists(sRootFolder & AddTrailingDelimiter(sNewFolder) & sNewFilename) Then
                            sTemp = "=HYPERLINK(""" + sNewLink + """, """ + sDisplay + """)"
                            oCurrent.Cells(iRow, iMediaCol + iColAdd).Formula = sTemp
                        End If
                    End If
                End If
            End If
            
            iColAdd = iColAdd + 1
        Wend
    Next iRow
    
    If ActiveWorkbook.ActiveSheet.Name <> oCurrent.Name Then
        oCurrent.Select
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Sub FormatColumns()
    Dim iCol As Long
    
    ForceFindExtents
    
    Call FormatColumnByName("Event.Start Clock", "dd/mm/yyyy HH:mm:ss")
    Call FormatColumnByName("Event.End Clock", "dd/mm/yyyy HH:mm:ss")
    
    Call FormatColumnByName("Start - Survey - Pipeline.KP", "0.0000")
    Call FormatColumnByName("End - Survey - Pipeline.KP", "0.0000")
    Call FormatColumnByName("Start - Survey - Standard.Depth", "0.0")
    Call FormatColumnByName("End - Survey - Standard.Depth", "0.0")
    
    
    iCol = Find_Column("Finding.Reason")
    If iCol > 0 Then
        Columns(iCol).ColumnWidth = 50
    End If
    
    iCol = Find_Column("Commentary.Notes")
    If iCol > 0 Then
        Columns(iCol).ColumnWidth = 50
    End If
    
    Cells.EntireRow.AutoFit
    
    Cells(2, 1).Select
End Sub
