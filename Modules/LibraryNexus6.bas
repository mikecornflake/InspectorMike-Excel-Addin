Attribute VB_Name = "LibraryNexus6"
Private Function AddRow(iRow As Long, AOriginal As String, ANew As String, Optional ADefault As String = "") As Integer
    Cells(iRow, 1).Value = AOriginal
    Cells(iRow, 2).Value = ANew
    Cells(iRow, 3).Value = ADefault
    AddRow = iRow + 1
End Function

Public Sub AddColumnNamesLookup()
    Dim oSheet As Worksheet
    Dim iRow As Long
    
    If Not Sheet_Exists("ColumnNames") Then
        ActiveWorkbook.Sheets.Add.Name = "ColumnNames"
    End If
    
    Set oSheet = FindSheet("ColumnNames")
    oSheet.Move After:=Sheets(Sheets.Count)

    oSheet.Activate
    
    '-----------  SURVEY FILE
    'Date Time
    'Easting
    'Northing
    'Kp
    'Dol
    'Heading
    'Pitch
    'Roll
    'CP Reading
    'TOP
    'BOP
    'LSB
    'RSB
    'Temperature
    'Salinity
    'Velocity
    'Depth
    'LSH
    'RSH
    'DVLDist
    
    If Cells(1, 1).Value = "" Then
        iRow = 1
        
        iRow = AddRow(iRow, "Original", "New", "Default Value")
        
        iRow = AddRow(iRow, "Date Time", "Survey Data.Clock")
        
        iRow = AddRow(iRow, "Easting", "Survey - Standard.Easting")
        iRow = AddRow(iRow, "Northing", "Survey - Standard.Northing")
        iRow = AddRow(iRow, "Depth", "Survey - Standard.Depth")
        iRow = AddRow(iRow, "LSH", "Survey - Standard.Elevation")
        
        iRow = AddRow(iRow, "Heading", "Other Fields.Heading")
        iRow = AddRow(iRow, "Temperature", "Other Fields.Temperature")
        iRow = AddRow(iRow, "CP reading", "Other Fields.Spare1")
        iRow = AddRow(iRow, "Pitch", "Other Fields.Spare2")
        iRow = AddRow(iRow, "Roll", "Other Fields.Spare3")
        iRow = AddRow(iRow, "Salinity", "Other Fields.Spare4")
        
        iRow = AddRow(iRow, "KP", "Survey - Pipeline.KP")
        iRow = AddRow(iRow, "DOL", "Survey - Pipeline.Offset")
        iRow = AddRow(iRow, "BOP", "Survey - Pipeline.BoP")
        iRow = AddRow(iRow, "TOP", "Survey - Pipeline.ToP")
        iRow = AddRow(iRow, "LSB", "Survey - Pipeline.Left")
        iRow = AddRow(iRow, "RSB", "Survey - Pipeline.Right")
        iRow = AddRow(iRow, "DVLDist", "Survey - Pipeline.Distance")
        
        iRow = AddRow(iRow, "Survey Data.Survey Set", "Survey Data.Survey Set", "2022 Q2 Processed")
        iRow = AddRow(iRow, "Event.Workpack", "Event.Workpack", "2022 Malampaya IRM P1")
        iRow = AddRow(iRow, "Asset Location.Full Location", "Asset Location.Full Location", "Shell Philippines Exploration / Malampaya / 24GEP OU / P/line SSIV - 100msw Contour / 24GEP")
        
        BasicTidy
    End If
End Sub

Public Sub RenameColumns(AToNew As Boolean)
    Dim oNames As Worksheet
    Dim oData As Worksheet
    Dim sOriginal As String
    Dim sNew As String
    Dim sDefault As String
    
    Dim iNameCount As Long
    Dim iRow As Long
    Dim iCol As Long

    Set oNames = FindSheet("ColumnNames")
    
    If IsNull(oSheet) Then
        MsgBox ("Tabsheet 'ColumnNames' not found")
    Else
        If ActiveSheet.Name = "ColumnNames" Then
            MsgBox ("Please switch to the TabSheet with data before running this routine")
        Else
            ' Yay, we can run
            Set oData = ActiveSheet
            
            oNames.Activate
            ForceFindExtents
            
            iNameCount = FLastRow
            
            oData.Activate
            ForceFindExtents
            
            ' Search over the table in oNames, but make changes in oData (which is selected & visible)
            For iRow = 2 To iNameCount
                If AToNew Then
                    sOriginal = oNames.Cells(iRow, 1).Value
                    sNew = oNames.Cells(iRow, 2).Value
                Else
                    sNew = oNames.Cells(iRow, 1).Value
                    sOriginal = oNames.Cells(iRow, 2).Value
                End If
                
                If Not Rename_Column(sOriginal, sNew) Then
                    ' If neither the old, nor the new exist, then add a new column called sNew
                    If (Find_Column(sOriginal) = -1) And (Find_Column(sNew) = -1) Then
                        iCol = Add_Column(sNew)
                        
                        sDefault = Trim(oNames.Cells(iRow, 3).Value)
                        
                        If sDefault <> "" Then
                            Cells(2, iCol).Value = sDefault
                            Cells(2, iCol).Select
                            Selection.Copy
                            Range(Cells(2, iCol), Cells(FLastRow, iCol)).Select
                            ActiveSheet.Paste
                        End If
                    End If
                End If
            Next iRow
            
            BasicTidy (False)
        End If
    End If
    
    frmRenameColumns.Hide
End Sub

Public Sub PrepareNexusImportFromCurrentSheet()
    If Sheet_Exists("Survey Import") Then
        MsgBox "Sheet called 'Survey Import' already exists"
        Exit Sub
    End If
    
    ' Rename this sheet to "Original"
    If Not Sheet_Exists("Original") Then
        If (ActiveSheet.Name <> "Original") And (ActiveSheet.Name <> "ColumnNames") And (ActiveSheet.Name <> "PL _ Profile") Then
            ActiveSheet.Name = "Original"
        End If
        ActiveSheet.Move Before:=Sheets(1)
    End If
    
    If Not Sheet_Exists("ColumnNames") Then
        AddColumnNamesLookup
        
        MsgBox "A sheet called 'ColumnNames' has just been added. " & vbCrLf & _
               "Please ensure the lookups and 'Survey Set' are correct before proceeding."
          
        Exit Sub
    End If
    
    If Sheet_Exists("Original") Then
        Sheets("Original").Select
    ElseIf ActiveSheet.Name = "ColumnNames" Then
        MsgBox "Please switch to the Excel worksheet with the original survey values first"
        
        Exit Sub
    End If
    
    If MsgBox("This will save the current file as an Excel workbook, " & vbCrLf & _
              "then rename & copy the current sheet. " & vbCrLf & vbCrLf & _
              "The new sheet will be processed to ensure a Survey " & vbCrLf & _
              "record exists at least every 3 seconds" & vbCrLf & vbCrLf & _
              "Are you sure you wish to continue", vbOKCancel) <> vbOK Then
        Exit Sub
    End If
      
    ' First save, this file as .xlsx
    SaveAsXLSX
    
    ActiveSheet.Copy After:=Sheets(1)
    Sheets(2).Select
    Sheets(2).Name = "Survey Import"
    
    EnforcePipelineDepthPolarity (False)
    
    Interpolate_Nav_To_3_Sec
    
    RenameColumns (True) ' This adds the missing columns
    
    ' But we don't want all the missing columns here
    Delete_Column ("Event.Workpack")
    Delete_Column ("Asset Location.Full Location")
    
    BasicTidyAndFormatColumns
    
    Sheets("Survey Import").Select
    ExportCurrentWorkSheetAsCSV
    
    Prepare_PL_Profile_Import
    
    Sheets("PL - Profile").Select
    ExportCurrentWorkSheetAsCSV
    
    MsgBox "Processing complete:" & vbCrLf & _
           "  - 'Survey Import' exported as CSV," & vbCrLf & _
           "  - 'PL - Profile' exported as CSV. " & vbCrLf & vbCrLf & _
           " CSV Files ready for importing into Nexus 6"
End Sub

Public Function EnforcePipelineDepthPolarity(ADepthPositive As Boolean)
    Dim iTOPCol As Long
    Dim iBOPCol As Long
    Dim iLSBCol As Long
    Dim iRSBCol As Long
    
    Dim iRow As Long
    Dim dTemp As Double
    Dim dTOP As Double
    Dim dBoP As Double
    Dim dLSB As Double
    Dim dRSB As Double
    
    ForceFindExtents
    
    iTOPCol = FindFirstColumn(Array("TOP", "Survey - Pipeline.ToP", "PL - Profile.Top of Pipe"))
    iBOPCol = FindFirstColumn(Array("BOP", "Survey - Pipeline.BoP", "PL - Profile.Bottom of Pipe"))
    iLSBCol = FindFirstColumn(Array("LSB", "Survey - Pipeline.Left", "PL - Profile.Left Seabed"))
    iRSBCol = FindFirstColumn(Array("RSB", "Survey - Pipeline.Right", "PL - Profile.Right Seabed"))
    
    For iRow = 2 To FLastRow
        Application.StatusBar = iRow & " of " & FLsatRow
        Cells(iRow, 1).Select
        
        dTOP = Abs(Cells(iRow, iTOPCol).Value)
        dBoP = Abs(Cells(iRow, iBOPCol).Value)
        dLSB = Abs(Cells(iRow, iLSBCol).Value)
        dRSB = Abs(Cells(iRow, iRSBCol).Value)
        
        If ADepthPositive Then
            Cells(iRow, iTOPCol).Value = min(dTOP, dBoP)
            Cells(iRow, iBOPCol).Value = max(dTOP, dBoP)
            
            Cells(iRow, iLSBCol).Value = dLSB
            Cells(iRow, iRSBCol).Value = dRSB
        Else
            dTOP = -1 * dTOP
            dBoP = -1 * dBoP
            
            Cells(iRow, iTOPCol).Value = max(dTOP, dBoP)
            Cells(iRow, iBOPCol).Value = min(dTOP, dBoP)
            
            Cells(iRow, iLSBCol).Value = -1 * dLSB
            Cells(iRow, iRSBCol).Value = -1 * dRSB
        End If
    Next iRow
    
    Cells(2, 1).Select
    Application.StatusBar = ""
End Function

Public Sub Prepare_PL_Profile_Import()
    Dim oSheet As Worksheet
    
    Dim iCol As Long
    Dim sTemp As String
    Dim oPLProfile As Worksheet
    Dim oOriginal As Worksheet
    Dim oColumnNames As Worksheet
    
    ' This works from the Original Data
    If Sheet_Exists("Original") Then
        Sheets("Original").Select
    End If
    
    Set oOriginal = ActiveSheet
    
    ActiveSheet.Move Before:=Sheets(1)
    ActiveSheet.Copy After:=Sheets(1)
    Sheets(2).Select
    Sheets(2).Name = "PL - Profile"
    Set oPLProfile = ActiveSheet
    
    Set oColumnNames = Sheets("ColumnNames")

    '----------Required Columns-------------------
    ' Event.Workpack
    ' Event.Event Type
    ' Asset Location.Full Location
    ' Event.Survey Set
    ' Event.Start Clock
    ' Event.End Clock
    ' PL - Profile.Top of Pipe
    ' PL - Profile.Bottom of Pipe
    ' PL - Profile.Left Seabed
    ' PL - Profile.Right Seabed

    ForceFindExtents
    
    iCol = FindFirstColumn(Array("TOP", "Survey - Pipeline.ToP", "PL - Profile.Top of Pipe"))
    Cells(1, iCol).Value = "PL - Profile.Top of Pipe"
    
    iCol = FindFirstColumn(Array("BOP", "Survey - Pipeline.BoP", "PL - Profile.Bottom of Pipe"))
    Cells(1, iCol).Value = "PL - Profile.Bottom of Pipe"
    
    iCol = FindFirstColumn(Array("LSB", "Survey - Pipeline.Left", "PL - Profile.Left Seabed"))
    Cells(1, iCol).Value = "PL - Profile.Left Seabed"
    
    iCol = FindFirstColumn(Array("RSB", "Survey - Pipeline.Right", "PL - Profile.Right Seabed"))
    Cells(1, iCol).Value = "PL - Profile.Right Seabed"
    
    iCol = FindFirstColumn(Array("Date Time", "Survey Data.Clock", "Event.Start Clock"))
    Cells(1, iCol).Value = "Event.Start Clock"
    
    Call Copy_Column("Event.Start Clock", "Event.End Clock")
    
    oPLProfile.Select
    ForceFindExtents ' used for Populate
    
    Ensure_Column ("Event.Workpack")
    oColumnNames.Select
    sTemp = Lookup("Original", "Event.Workpack", "Default Value")
    oPLProfile.Select
    Call PopulateColumn("Event.Workpack", sTemp)
    
    Ensure_Column ("Asset Location.Full Location")
    oColumnNames.Select
    sTemp = Lookup("Original", "Asset Location.Full Location", "Default Value")
    oPLProfile.Select
    Call PopulateColumn("Asset Location.Full Location", sTemp)
    
    Ensure_Column ("Event.Workpack")
    oColumnNames.Select
    sTemp = Lookup("Original", "Event.Workpack", "Default Value")
    oPLProfile.Select
    Call PopulateColumn("Event.Workpack", sTemp)
    
    Ensure_Column ("Event.Survey Set")
    oColumnNames.Select
    sTemp = Lookup("Original", "Survey Data.Survey Set", "Default Value")
    If sTemp = "" Then
        sTemp = Lookup("Original", "Event.Survey Set", "Default Value")
    End If
    oPLProfile.Select
    Call PopulateColumn("Event.Survey Set", sTemp)
    
    Ensure_Column ("Event.Event Type")
    Call PopulateColumn("Event.Event Type", "PL - Profile")

    Call Move_Column("Event.Workpack", 1)
    Call Move_Column("Event.Event Type", 2)
    Call Move_Column("Asset Location.Full Location", 3)
    Call Move_Column("Event.Survey Set", 4)
    Call Move_Column("Event.Start Clock", 5)
    Call Move_Column("Event.End Clock", 6)
    Call Move_Column("PL - Profile.Top of Pipe", 7)
    Call Move_Column("PL - Profile.Bottom of Pipe", 8)
    Call Move_Column("PL - Profile.Left Seabed", 9)
    Call Move_Column("PL - Profile.Right Seabed", 10)
    
    Call Delete_Column("Easting")
    Call Delete_Column("Northing")
    Call Delete_Column("Kp")
    Call Delete_Column("Dol")
    Call Delete_Column("Heading")
    Call Delete_Column("Pitch")
    Call Delete_Column("Roll")
    Call Delete_Column("CP Reading")
    Call Delete_Column("Temperature")
    Call Delete_Column("Salinity")
    Call Delete_Column("Velocity")
    Call Delete_Column("Depth")
    Call Delete_Column("LSH")
    Call Delete_Column("RSH")
    Call Delete_Column("DVLDist")
    
    Call Delete_Column("Survey - Standard.Easting")
    Call Delete_Column("Survey - Standard.Northing")
    Call Delete_Column("Survey - Pipeline.KP")
    Call Delete_Column("Survey - Pipeline.Offset")
    Call Delete_Column("Other Fields.Heading")
    Call Delete_Column("Other Fields.Spare1")
    Call Delete_Column("Other Fields.Spare2")
    Call Delete_Column("Other Fields.Spare3")
    Call Delete_Column("Other Fields.Spare4")
    Call Delete_Column("Other Fields.Temperature")
    Call Delete_Column("Survey - Standard.Depth")
    Call Delete_Column("Survey - Standard.Elevation")
    Call Delete_Column("Survey - Pipeline.Distance")
    Call Delete_Column("Survey Data.Survey Set")
    
    EnforcePipelineDepthPolarity (False)
    BasicTidyAndFormatColumns
End Sub
