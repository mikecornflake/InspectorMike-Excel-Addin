Attribute VB_Name = "tstArray"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass
'            - Tests and corrections by copilot/ChatGPT-5

Option Private Module

Public Sub Test_LibraryArray()
    ActiveTestModule = "libArray"

    ' ===== Array_IndexOf tests =====
    Dim arr1 As Variant
    arr1 = Array("apple", "banana", "cherry")

    Call AssertEqual("Array_IndexOf - found", 1, Array_IndexOf(arr1, "banana"))
    Call AssertEqual("Array_IndexOf - not found", -1, Array_IndexOf(arr1, "grape"))
    Call AssertEqual("Array_IndexOf - numeric", 0, Array_IndexOf(Array(10, 20, 30), 10))
    Call AssertEqual("Array_IndexOf - empty array", -1, Array_IndexOf(Array(), "x"))

    ' ===== Collection_IndexOf - collection tests =====
    Dim colTest As New Collection
    colTest.Add "apple"
    colTest.Add "banana"
    colTest.Add "cherry"
    
    Call AssertEqual("Collection_IndexOf - collection found", 2, Collection_IndexOf(colTest, "banana"))
    Call AssertEqual("Collection_IndexOf - collection not found", -1, Collection_IndexOf(colTest, "grape"))

    ' ===== Collection_Sort tests =====
    Dim colAsc As New Collection
    colAsc.Add 5
    colAsc.Add 2
    colAsc.Add 9
    colAsc.Add 1

    Dim sortedAsc As Collection
    Set sortedAsc = Collection_Sort(colAsc, True)

    Call AssertEqual("Collection_Sort Asc - first", 1, sortedAsc(1))
    Call AssertEqual("Collection_Sort Asc - last", 9, sortedAsc(sortedAsc.Count))

    Dim colDesc As New Collection
    colDesc.Add "b"
    colDesc.Add "d"
    colDesc.Add "a"

    Dim sortedDesc As Collection
    Set sortedDesc = Collection_Sort(colDesc, False)

    Call AssertEqual("Collection_Sort Desc - first", "d", sortedDesc(1))
    Call AssertEqual("Collection_Sort Desc - last", "a", sortedDesc(sortedDesc.Count))

    ' ===== Edge case: single item =====
    Dim colSingle As New Collection
    colSingle.Add 42
    Set colSingle = Collection_Sort(colSingle)
    Call AssertEqual("Collection_Sort - single item", 42, colSingle(1))

    ' ===== Edge case: empty collection =====
    Dim colEmpty As New Collection
    Set colEmpty = Collection_Sort(colEmpty)
    Call AssertEqual("Collection_Sort - empty collection", 0, colEmpty.Count)

    ' ===== Array_Sort tests =====
    Dim arrAsc As Variant
    arrAsc = Array(5, 2, 9, 1)
    Dim sortedArrAsc As Variant
    sortedArrAsc = Array_Sort(arrAsc, True)
    Call AssertEqual("Array_Sort Asc - first", 1, sortedArrAsc(0))
    Call AssertEqual("Array_Sort Asc - last", 9, sortedArrAsc(UBound(sortedArrAsc)))

    Dim arrDesc As Variant
    arrDesc = Array("b", "d", "a")
    Dim sortedArrDesc As Variant
    sortedArrDesc = Array_Sort(arrDesc, False)
    Call AssertEqual("Array_Sort Desc - first", "d", sortedArrDesc(0))
    Call AssertEqual("Array_Sort Desc - last", "a", sortedArrDesc(UBound(sortedArrDesc)))

    ' ===== Edge case: single item =====
    Dim arrSingle As Variant
    arrSingle = Array(42)
    Dim sortedSingle As Variant
    sortedSingle = Array_Sort(arrSingle)
    Call AssertEqual("Array_Sort - single item", 42, sortedSingle(0))

    ' ===== Edge case: empty array =====
    Dim arrEmpty As Variant
    arrEmpty = Array()
    Dim sortedEmpty As Variant
    sortedEmpty = Array_Sort(arrEmpty)
    Call AssertEqual("Array_Sort - empty array length", -1, UBound(sortedEmpty) - LBound(sortedEmpty))

    ' ===== Already sorted array =====
    Dim arrSorted As Variant
    arrSorted = Array(1, 2, 3, 4)
    Dim resultSorted As Variant
    resultSorted = Array_Sort(arrSorted, True)
    Call AssertEqual("Array_Sort - already sorted", 2, resultSorted(1))

    ' ===== All equal elements =====
    Dim arrEqual As Variant
    arrEqual = Array(7, 7, 7, 7)
    Dim sortedEqual As Variant
    sortedEqual = Array_Sort(arrEqual)
    Call AssertEqual("Array_Sort - all equal", 7, sortedEqual(2))
    
    ' ===== Collection_Sort tests - collection of array(key, payload) =====
    Dim colKeyedAsc As New Collection
    colKeyedAsc.Add Array(3, "Charlie")
    colKeyedAsc.Add Array(1, "Alpha")
    colKeyedAsc.Add Array(2, "Bravo")

    Dim sortedKeyedAsc As Collection
    Set sortedKeyedAsc = Collection_Sort(colKeyedAsc, True)

    Call AssertEqual("Collection_Sort keyed Asc - first key", 1, sortedKeyedAsc(1)(0))
    Call AssertEqual("Collection_Sort keyed Asc - first payload", "Alpha", sortedKeyedAsc(1)(1))
    Call AssertEqual("Collection_Sort keyed Asc - last key", 3, sortedKeyedAsc(sortedKeyedAsc.Count)(0))
    Call AssertEqual("Collection_Sort keyed Asc - last payload", "Charlie", sortedKeyedAsc(sortedKeyedAsc.Count)(1))

    Dim colKeyedDesc As New Collection
    colKeyedDesc.Add Array(3, "Charlie")
    colKeyedDesc.Add Array(1, "Alpha")
    colKeyedDesc.Add Array(2, "Bravo")

    Dim sortedKeyedDesc As Collection
    Set sortedKeyedDesc = Collection_Sort(colKeyedDesc, False)

    Call AssertEqual("Collection_Sort keyed Desc - first key", 3, sortedKeyedDesc(1)(0))
    Call AssertEqual("Collection_Sort keyed Desc - first payload", "Charlie", sortedKeyedDesc(1)(1))
    Call AssertEqual("Collection_Sort keyed Desc - last key", 1, sortedKeyedDesc(sortedKeyedDesc.Count)(0))
    Call AssertEqual("Collection_Sort keyed Desc - last payload", "Alpha", sortedKeyedDesc(sortedKeyedDesc.Count)(1))

    ' ===== Collection_Sort tests - duplicate keys =====
    Dim colKeyedDup As New Collection
    colKeyedDup.Add Array(2, "Bravo1")
    colKeyedDup.Add Array(1, "Alpha")
    colKeyedDup.Add Array(2, "Bravo2")

    Dim sortedKeyedDup As Collection
    Set sortedKeyedDup = Collection_Sort(colKeyedDup, True)

    Call AssertEqual("Collection_Sort keyed Dup - first key", 1, sortedKeyedDup(1)(0))
    Call AssertEqual("Collection_Sort keyed Dup - first payload", "Alpha", sortedKeyedDup(1)(1))
    Call AssertEqual("Collection_Sort keyed Dup - second key", 2, sortedKeyedDup(2)(0))
    Call AssertEqual("Collection_Sort keyed Dup - third key", 2, sortedKeyedDup(3)(0))

    ' ===== Array_Sort tests - array of array(key, payload) =====
    Dim arrKeyedAsc As Variant
    arrKeyedAsc = Array( _
        Array(3, "Charlie"), _
        Array(1, "Alpha"), _
        Array(2, "Bravo") _
    )

    Dim sortedArrKeyedAsc As Variant
    sortedArrKeyedAsc = Array_Sort(arrKeyedAsc, True)

    Call AssertEqual("Array_Sort keyed Asc - first key", 1, sortedArrKeyedAsc(0)(0))
    Call AssertEqual("Array_Sort keyed Asc - first payload", "Alpha", sortedArrKeyedAsc(0)(1))
    Call AssertEqual("Array_Sort keyed Asc - last key", 3, sortedArrKeyedAsc(UBound(sortedArrKeyedAsc))(0))
    Call AssertEqual("Array_Sort keyed Asc - last payload", "Charlie", sortedArrKeyedAsc(UBound(sortedArrKeyedAsc))(1))

    Dim arrKeyedDesc As Variant
    arrKeyedDesc = Array( _
        Array(3, "Charlie"), _
        Array(1, "Alpha"), _
        Array(2, "Bravo") _
    )

    Dim sortedArrKeyedDesc As Variant
    sortedArrKeyedDesc = Array_Sort(arrKeyedDesc, False)

    Call AssertEqual("Array_Sort keyed Desc - first key", 3, sortedArrKeyedDesc(0)(0))
    Call AssertEqual("Array_Sort keyed Desc - first payload", "Charlie", sortedArrKeyedDesc(0)(1))
    Call AssertEqual("Array_Sort keyed Desc - last key", 1, sortedArrKeyedDesc(UBound(sortedArrKeyedDesc))(0))
    Call AssertEqual("Array_Sort keyed Desc - last payload", "Alpha", sortedArrKeyedDesc(UBound(sortedArrKeyedDesc))(1))

    ' ===== Array_Sort tests - single keyed item =====
    Dim arrKeyedSingle As Variant
    arrKeyedSingle = Array(Array(42, "Only"))
    
    Dim sortedArrKeyedSingle As Variant
    sortedArrKeyedSingle = Array_Sort(arrKeyedSingle, True)
    
    Call AssertEqual("Array_Sort keyed Single - key", 42, sortedArrKeyedSingle(0)(0))
    Call AssertEqual("Array_Sort keyed Single - payload", "Only", sortedArrKeyedSingle(0)(1))
End Sub

