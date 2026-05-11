Attribute VB_Name = "tstTable"
Option Explicit
Option Private Module

Public Sub Test_libTable()
    ActiveTestModule = "libTable"

    Call Test_Table_FindColumn
    Call Test_Table_FindFirstColumn
    Call Test_Table_FindInColumn
    Call Test_Table_TableRange
    Call Test_Table_FindTableExtents
    Call Test_Table_CompareRows
    Call Test_Table_SortTable
    Call Test_Table_DeleteDuplicateRows
    Call Test_Table_DeleteDuplicateRowsByColumn
    Call Test_Table_EnsureInsertAppendDeleteRenameColumn
    Call Test_Table_GetColumnLetter
    Call Test_Table_IsRowPopulated
    Call Test_Table_FormatColumn
    Call Test_Table_ConvertColumnToValues
    Call Test_Table_LookupColumnValue
    Call Test_Table_PopulateColumn
    Call Test_Table_MoveColumn
    Call Test_Table_MoveColumnToName
    Call Test_Table_CopyColumn
End Sub

Private Sub Test_Table_FindColumn()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_FindColumn")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Depth"
    ws.Cells(1, 3).Value = "Comment"

    Call AssertEqual("FindColumn - existing", 2, FindColumn(ws, "Depth"))
    Call AssertEqual("FindColumn - case insensitive", 2, FindColumn(ws, "depth"))
    Call AssertEqual("FindColumn - missing", -1, FindColumn(ws, "Missing"))

    Call DeleteTestSheet("test_Table_FindColumn")
End Sub

Private Sub Test_Table_FindFirstColumn()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_FindFirstColumn")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Depth (m)"
    ws.Cells(1, 3).Value = "Comment"

    Call AssertEqual("FindFirstColumn - finds second option", 2, FindFirstColumn(ws, Array("Depth", "Depth (m)")))
    Call AssertEqual("FindFirstColumn - missing", -1, FindFirstColumn(ws, Array("Missing1", "Missing2")))

    Call DeleteTestSheet("test_Table_FindFirstColumn")
End Sub

Private Sub Test_Table_FindInColumn()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_FindInColumn")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(2, 1).Value = "A001"
    ws.Cells(3, 1).Value = "A002"
    ws.Cells(4, 1).Value = "A003"

    Call AssertEqual("FindInColumn - existing", 3, FindInColumn(ws, 1, "A002"))
    Call AssertEqual("FindInColumn - missing", 0, FindInColumn(ws, 1, "A999"))

    Call DeleteTestSheet("test_Table_FindInColumn")
End Sub

Private Sub Test_Table_TableRange()
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = CreateTestSheet("test_Table_TableRange")

    ws.Cells(1, 1).Value = "A"
    ws.Cells(1, 3).Value = "C"
    ws.Cells(5, 3).Value = "X"

    Set rng = TableRange(ws)

    Call AssertEqual("TableRange - address", "$A$1:$C$5", rng.Address)

    Set rng = TableRange(ws, "B2", 4, 3)

    Call AssertEqual("TableRange - explicit address", "$B$2:$C$4", rng.Address)

    Call DeleteTestSheet("test_Table_TableRange")
End Sub

Private Sub Test_Table_FindTableExtents()
    Dim ws As Worksheet
    Dim lLastRow As Long
    Dim lLastCol As Long

    Set ws = CreateTestSheet("test_Table_FindTableExtents")

    ws.Cells(1, 1).Value = "A"
    ws.Cells(6, 4).Value = "X"

    Call FindTableExtents(ws, lLastRow, lLastCol)

    Call AssertEqual("FindTableExtents - last row", 6, lLastRow)
    Call AssertEqual("FindTableExtents - last column", 4, lLastCol)

    Call DeleteTestSheet("test_Table_FindTableExtents")
End Sub

Private Sub Test_Table_CompareRows()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_CompareRows")

    ws.Cells(1, 1).Value = "A"
    ws.Cells(1, 2).Value = "B"

    ws.Cells(2, 1).Value = "X"
    ws.Cells(2, 2).Value = "Y"

    ws.Cells(3, 1).Value = "X"
    ws.Cells(3, 2).Value = "Y"

    ws.Cells(4, 1).Value = "X"
    ws.Cells(4, 2).Value = "Z"

    Call AssertTrue("CompareRows - same", CompareRows(ws, 2, 3, 2))
    Call AssertFalse("CompareRows - different", CompareRows(ws, 2, 4, 2))

    Call DeleteTestSheet("test_Table_CompareRows")
End Sub

Private Sub Test_Table_SortTable()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_SortTable")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Seq"

    ws.Cells(2, 1).Value = "B"
    ws.Cells(2, 2).Value = 2

    ws.Cells(3, 1).Value = "A"
    ws.Cells(3, 2).Value = 1

    Call SortTable(ws, 1)

    Call AssertEqual("SortTable - row 2 asset", "A", ws.Cells(2, 1).Value)
    Call AssertEqual("SortTable - row 3 asset", "B", ws.Cells(3, 1).Value)

    Call DeleteTestSheet("test_Table_SortTable")
End Sub

Private Sub Test_Table_DeleteDuplicateRows()
    Dim ws As Worksheet
    Dim lDeleted As Long

    Set ws = CreateTestSheet("test_Table_DeleteDuplicateRows")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Status"

    ws.Cells(2, 1).Value = "A"
    ws.Cells(2, 2).Value = "Open"

    ws.Cells(3, 1).Value = "A"
    ws.Cells(3, 2).Value = "Open"

    ws.Cells(4, 1).Value = "B"
    ws.Cells(4, 2).Value = "Closed"

    lDeleted = DeleteDuplicateRows(ws)

    Call AssertEqual("DeleteDuplicateRows - count", 1, lDeleted)
    Call AssertEqual("DeleteDuplicateRows - last row", 3, LastUsedRow(ws))

    Call DeleteTestSheet("test_Table_DeleteDuplicateRows")
End Sub

Private Sub Test_Table_DeleteDuplicateRowsByColumn()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_DeleteDuplicateRowsByColumn")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Status"

    ws.Cells(2, 1).Value = "A"
    ws.Cells(2, 2).Value = "Open"

    ws.Cells(3, 1).Value = "A"
    ws.Cells(3, 2).Value = "Changed"

    ws.Cells(4, 1).Value = "B"
    ws.Cells(4, 2).Value = "Closed"

    Call DeleteDuplicateRowsByColumn(ws, 1)

    Call AssertEqual("DeleteDuplicateRowsByColumn - last row", 3, LastUsedRow(ws))
    Call AssertEqual("DeleteDuplicateRowsByColumn - kept first duplicate", "Open", ws.Cells(2, 2).Value)

    Call DeleteTestSheet("test_Table_DeleteDuplicateRowsByColumn")
End Sub

Private Sub Test_Table_EnsureInsertAppendDeleteRenameColumn()
    Dim ws As Worksheet
    Dim iCol As Long
    Dim bResult As Boolean

    Set ws = CreateTestSheet("test_Table_Columns")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Depth"

    iCol = EnsureColumn(ws, "Comment")
    Call AssertEqual("EnsureColumn - created", 3, iCol)

    iCol = EnsureColumn(ws, "Depth")
    Call AssertEqual("EnsureColumn - existing", 2, iCol)

    iCol = InsertColumn(ws, "Date", 2)
    Call AssertEqual("InsertColumn - inserted", 2, iCol)
    Call AssertEqual("InsertColumn - header", "Date", ws.Cells(1, 2).Value)

    iCol = AppendColumn(ws, "Inspector")
    Call AssertEqual("AppendColumn - appended", 5, iCol)
    Call AssertEqual("AppendColumn - header", "Inspector", ws.Cells(1, 5).Value)

    bResult = RenameColumn(ws, "Inspector", "User")
    Call AssertTrue("RenameColumn - true", bResult)
    Call AssertEqual("RenameColumn - header", "User", ws.Cells(1, 5).Value)

    bResult = DeleteColumn(ws, "Date")
    Call AssertTrue("DeleteColumn - true", bResult)
    Call AssertEqual("DeleteColumn - shifted left", "Depth", ws.Cells(1, 2).Value)

    Call DeleteTestSheet("test_Table_Columns")
End Sub

Private Sub Test_Table_GetColumnLetter()
    Call AssertEqual("GetColumnLetter - A", "A", GetColumnLetter(1))
    Call AssertEqual("GetColumnLetter - Z", "Z", GetColumnLetter(26))
    Call AssertEqual("GetColumnLetter - AA", "AA", GetColumnLetter(27))
    Call AssertEqual("GetColumnLetter - AB", "AB", GetColumnLetter(28))
    Call AssertEqual("GetColumnLetter - AZ", "AZ", GetColumnLetter(52))
    Call AssertEqual("GetColumnLetter - BA", "BA", GetColumnLetter(53))
End Sub

Private Sub Test_Table_IsRowPopulated()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_IsRowPopulated")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Depth"

    ws.Cells(2, 1).Value = ""
    ws.Cells(2, 2).Value = ""

    ws.Cells(3, 2).Value = "X"

    Call AssertFalse("IsRowPopulated - empty row", IsRowPopulated(ws, 2))
    Call AssertTrue("IsRowPopulated - populated row", IsRowPopulated(ws, 3))
    Call AssertFalse("IsRowPopulated - header ignored", IsRowPopulated(ws, 1))

    Call DeleteTestSheet("test_Table_IsRowPopulated")
End Sub

Private Sub Test_Table_FormatColumn()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_FormatColumn")

    ws.Cells(1, 1).Value = "Date"
    ws.Cells(1, 2).Value = "Depth"

    Call FormatColumn(ws, 2, "0.00")
    Call AssertEqual("FormatColumn - direct", "0.00", ws.Columns(2).NumberFormat)

    Call FormatColumnByName(ws, "Date", "yyyy-mm-dd")
    Call AssertEqual("FormatColumnByName - by name", "yyyy-mm-dd", ws.Columns(1).NumberFormat)

    Call FormatColumnByNames(ws, Array("Depth", "Missing"), "0.000")
    Call AssertEqual("FormatColumnByNames - array", "0.000", ws.Columns(2).NumberFormat)

    Call DeleteTestSheet("test_Table_FormatColumn")
End Sub

Private Sub Test_Table_ConvertColumnToValues()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_ConvertColumnToValues")

    ws.Cells(1, 1).Value = "Value"
    ws.Cells(2, 1).Formula = "=1+2"

    Call ConvertColumnToValues(ws, "Value")

    Call AssertEqual("ConvertColumnToValues - value", 3, ws.Cells(2, 1).Value)
    Call AssertFalse("ConvertColumnToValues - formula removed", ws.Cells(2, 1).HasFormula)

    Call DeleteTestSheet("test_Table_ConvertColumnToValues")
End Sub

Private Sub Test_Table_LookupColumnValue()
    Dim ws As Worksheet

    Set ws = CreateTestSheet("test_Table_LookupColumnValue")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Status"

    ws.Cells(2, 1).Value = "A001"
    ws.Cells(2, 2).Value = "Open"

    ws.Cells(3, 1).Value = "A002"
    ws.Cells(3, 2).Value = "Closed"

    Call AssertEqual("LookupColumnValue - found", "Closed", LookupColumnValue(ws, "Asset", "A002", "Status"))
    Call AssertEqual("LookupColumnValue - case insensitive", "Closed", LookupColumnValue(ws, "Asset", "a002", "Status"))
    Call AssertEqual("LookupColumnValue - missing value", "", LookupColumnValue(ws, "Asset", "A999", "Status"))
    Call AssertEqual("LookupColumnValue - missing lookup column", "", LookupColumnValue(ws, "Missing", "A002", "Status"))
    Call AssertEqual("LookupColumnValue - missing return column", "", LookupColumnValue(ws, "Asset", "A002", "Missing"))

    Call DeleteTestSheet("test_Table_LookupColumnValue")
End Sub

Private Sub Test_Table_PopulateColumn()
    Dim ws As Worksheet
    Dim bResult As Boolean

    Set ws = CreateTestSheet("test_Table_PopulateColumn")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(1, 2).Value = "Flag"

    ws.Cells(2, 1).Value = "A001"
    ws.Cells(3, 1).Value = "A002"

    bResult = PopulateColumn(ws, "Flag", "Y")

    Call AssertTrue("PopulateColumn - returns true", bResult)
    Call AssertEqual("PopulateColumn - row 2", "Y", ws.Cells(2, 2).Value)
    Call AssertEqual("PopulateColumn - row 3", "Y", ws.Cells(3, 2).Value)

    bResult = PopulateColumn(ws, "Missing", "Y")
    Call AssertFalse("PopulateColumn - missing column", bResult)

    Call DeleteTestSheet("test_Table_PopulateColumn")
End Sub

Private Sub Test_Table_MoveColumn()
    Dim ws As Worksheet
    Dim bResult As Boolean

    Set ws = CreateTestSheet("test_Table_MoveColumn")

    ws.Cells(1, 1).Value = "A"
    ws.Cells(1, 2).Value = "B"
    ws.Cells(1, 3).Value = "C"

    bResult = MoveColumn(ws, "A", 3)

    Call AssertTrue("MoveColumn - returns true", bResult)
    Call AssertEqual("MoveColumn - col 1", "B", ws.Cells(1, 1).Value)
    Call AssertEqual("MoveColumn - col 2", "A", ws.Cells(1, 2).Value)
    Call AssertEqual("MoveColumn - col 3", "C", ws.Cells(1, 3).Value)

    bResult = MoveColumn(ws, "Missing", 2)
    Call AssertFalse("MoveColumn - missing column", bResult)

    Call DeleteTestSheet("test_Table_MoveColumn")
End Sub

Private Sub Test_Table_MoveColumnToName()
    Dim ws As Worksheet
    Dim bResult As Boolean

    Set ws = CreateTestSheet("test_Table_MoveColumnToName")

    ws.Cells(1, 1).Value = "A"
    ws.Cells(1, 2).Value = "B"
    ws.Cells(1, 3).Value = "C"

    bResult = MoveColumnToName(ws, "A", "C")

    Call AssertTrue("MoveColumnToName - returns true", bResult)
    Call AssertEqual("MoveColumnToName - col 1", "B", ws.Cells(1, 1).Value)
    Call AssertEqual("MoveColumnToName - col 2", "A", ws.Cells(1, 2).Value)
    Call AssertEqual("MoveColumnToName - col 3", "C", ws.Cells(1, 3).Value)

    Call DeleteTestSheet("test_Table_MoveColumnToName")
End Sub

Private Sub Test_Table_CopyColumn()
    Dim ws As Worksheet
    Dim bResult As Boolean

    Set ws = CreateTestSheet("test_Table_CopyColumn")

    ws.Cells(1, 1).Value = "Asset"
    ws.Cells(2, 1).Value = "A001"
    ws.Cells(3, 1).Value = "A002"

    ws.Cells(1, 2).Value = "Status"
    ws.Cells(2, 2).Value = "Open"
    ws.Cells(3, 2).Value = "Closed"

    bResult = CopyColumn(ws, "Asset", "AssetCopy")

    Call AssertTrue("CopyColumn - returns true", bResult)
    Call AssertEqual("CopyColumn - new header", "AssetCopy", ws.Cells(1, 2).Value)
    Call AssertEqual("CopyColumn - copied row 2", "A001", ws.Cells(2, 2).Value)
    Call AssertEqual("CopyColumn - shifted old col 2", "Status", ws.Cells(1, 3).Value)

    bResult = CopyColumn(ws, "Missing", "MissingCopy")
    Call AssertFalse("CopyColumn - missing source", bResult)

    Call DeleteTestSheet("test_Table_CopyColumn")
End Sub
