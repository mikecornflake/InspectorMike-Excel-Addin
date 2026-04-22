Attribute VB_Name = "libArray"
' 2025 08 15 - Tests added.  Err, then routines fixed.  All tests now pass
'            - Tests and corrections by copilot/ChatGPT-5
'            - Added ArrayBubbleSort
'            - This unit handles both Collections and Arrays, they have different indexing rules (array 0 based, collections 1 based).
'            -     IndexOf modified to handle the differences

Option Explicit

Private Const INSERTIONSORT_THRESHOLD As Long = 7

' Can't do this in a procedure, but leaving code here so I don't have to google how to increase array again...
'
'Function Add(ByVal AValue As Variant, ByRef Aarr As Variant) As Integer
'    ReDim Preserve Aarr(UBound(Aarr) + 1)
'    Aarr(UBound(Aarr)) = AValue
'    Add = UBound(Aarr)
'End Function


Public Sub Test_LibraryArray()
    ActiveTestModule = "libArray"

    ' ===== IndexOf tests =====
    Dim arr1 As Variant
    arr1 = Array("apple", "banana", "cherry")

    Call AssertEqual("IndexOf - found", 1, IndexOf("banana", arr1))
    Call AssertEqual("IndexOf - not found", -1, IndexOf("grape", arr1))
    Call AssertEqual("IndexOf - numeric", 0, IndexOf(10, Array(10, 20, 30)))
    Call AssertEqual("IndexOf - empty array", -1, IndexOf("x", Array()))

    ' ===== IndexOf - collection tests =====
    Dim colTest As New Collection
    colTest.Add "apple"
    colTest.Add "banana"
    colTest.Add "cherry"
    
    Call AssertEqual("IndexOf - collection found", 2, IndexOf("banana", colTest))
    Call AssertEqual("IndexOf - collection not found", -1, IndexOf("grape", colTest))

    ' ===== CollectionBubbleSort tests =====
    Dim colAsc As New Collection
    colAsc.Add 5
    colAsc.Add 2
    colAsc.Add 9
    colAsc.Add 1

    Dim sortedAsc As Collection
    Set sortedAsc = CollectionBubbleSort(colAsc, True)

    Call AssertEqual("CollectionBubbleSort Asc - first", 1, sortedAsc(1))
    Call AssertEqual("CollectionBubbleSort Asc - last", 9, sortedAsc(sortedAsc.Count))

    Dim colDesc As New Collection
    colDesc.Add "b"
    colDesc.Add "d"
    colDesc.Add "a"

    Dim sortedDesc As Collection
    Set sortedDesc = CollectionBubbleSort(colDesc, False)

    Call AssertEqual("CollectionBubbleSort Desc - first", "d", sortedDesc(1))
    Call AssertEqual("CollectionBubbleSort Desc - last", "a", sortedDesc(sortedDesc.Count))

    ' ===== Edge case: single item =====
    Dim colSingle As New Collection
    colSingle.Add 42
    Set colSingle = CollectionBubbleSort(colSingle)
    Call AssertEqual("CollectionBubbleSort - single item", 42, colSingle(1))

    ' ===== Edge case: empty collection =====
    Dim colEmpty As New Collection
    Set colEmpty = CollectionBubbleSort(colEmpty)
    Call AssertEqual("CollectionBubbleSort - empty collection", 0, colEmpty.Count)

    ' ===== ArrayBubbleSort tests =====
    Dim arrAsc As Variant
    arrAsc = Array(5, 2, 9, 1)
    Dim sortedArrAsc As Variant
    sortedArrAsc = ArrayBubbleSort(arrAsc, True)
    Call AssertEqual("ArrayBubbleSort Asc - first", 1, sortedArrAsc(0))
    Call AssertEqual("ArrayBubbleSort Asc - last", 9, sortedArrAsc(UBound(sortedArrAsc)))

    Dim arrDesc As Variant
    arrDesc = Array("b", "d", "a")
    Dim sortedArrDesc As Variant
    sortedArrDesc = ArrayBubbleSort(arrDesc, False)
    Call AssertEqual("ArrayBubbleSort Desc - first", "d", sortedArrDesc(0))
    Call AssertEqual("ArrayBubbleSort Desc - last", "a", sortedArrDesc(UBound(sortedArrDesc)))

    ' ===== Edge case: single item =====
    Dim arrSingle As Variant
    arrSingle = Array(42)
    Dim sortedSingle As Variant
    sortedSingle = ArrayBubbleSort(arrSingle)
    Call AssertEqual("ArrayBubbleSort - single item", 42, sortedSingle(0))

    ' ===== Edge case: empty array =====
    Dim arrEmpty As Variant
    arrEmpty = Array()
    Dim sortedEmpty As Variant
    sortedEmpty = ArrayBubbleSort(arrEmpty)
    Call AssertEqual("ArrayBubbleSort - empty array length", -1, UBound(sortedEmpty) - LBound(sortedEmpty))

    ' ===== Already sorted array =====
    Dim arrSorted As Variant
    arrSorted = Array(1, 2, 3, 4)
    Dim resultSorted As Variant
    resultSorted = ArrayBubbleSort(arrSorted, True)
    Call AssertEqual("ArrayBubbleSort - already sorted", 2, resultSorted(1))

    ' ===== All equal elements =====
    Dim arrEqual As Variant
    arrEqual = Array(7, 7, 7, 7)
    Dim sortedEqual As Variant
    sortedEqual = ArrayBubbleSort(arrEqual)
    Call AssertEqual("ArrayBubbleSort - all equal", 7, sortedEqual(2))
    
    ' ===== CollectionBubbleSort tests - collection of array(key, payload) =====
    Dim colKeyedAsc As New Collection
    colKeyedAsc.Add Array(3, "Charlie")
    colKeyedAsc.Add Array(1, "Alpha")
    colKeyedAsc.Add Array(2, "Bravo")

    Dim sortedKeyedAsc As Collection
    Set sortedKeyedAsc = CollectionBubbleSort(colKeyedAsc, True)

    Call AssertEqual("CollectionBubbleSort keyed Asc - first key", 1, sortedKeyedAsc(1)(0))
    Call AssertEqual("CollectionBubbleSort keyed Asc - first payload", "Alpha", sortedKeyedAsc(1)(1))
    Call AssertEqual("CollectionBubbleSort keyed Asc - last key", 3, sortedKeyedAsc(sortedKeyedAsc.Count)(0))
    Call AssertEqual("CollectionBubbleSort keyed Asc - last payload", "Charlie", sortedKeyedAsc(sortedKeyedAsc.Count)(1))

    Dim colKeyedDesc As New Collection
    colKeyedDesc.Add Array(3, "Charlie")
    colKeyedDesc.Add Array(1, "Alpha")
    colKeyedDesc.Add Array(2, "Bravo")

    Dim sortedKeyedDesc As Collection
    Set sortedKeyedDesc = CollectionBubbleSort(colKeyedDesc, False)

    Call AssertEqual("CollectionBubbleSort keyed Desc - first key", 3, sortedKeyedDesc(1)(0))
    Call AssertEqual("CollectionBubbleSort keyed Desc - first payload", "Charlie", sortedKeyedDesc(1)(1))
    Call AssertEqual("CollectionBubbleSort keyed Desc - last key", 1, sortedKeyedDesc(sortedKeyedDesc.Count)(0))
    Call AssertEqual("CollectionBubbleSort keyed Desc - last payload", "Alpha", sortedKeyedDesc(sortedKeyedDesc.Count)(1))

    ' ===== CollectionBubbleSort tests - duplicate keys =====
    Dim colKeyedDup As New Collection
    colKeyedDup.Add Array(2, "Bravo1")
    colKeyedDup.Add Array(1, "Alpha")
    colKeyedDup.Add Array(2, "Bravo2")

    Dim sortedKeyedDup As Collection
    Set sortedKeyedDup = CollectionBubbleSort(colKeyedDup, True)

    Call AssertEqual("CollectionBubbleSort keyed Dup - first key", 1, sortedKeyedDup(1)(0))
    Call AssertEqual("CollectionBubbleSort keyed Dup - first payload", "Alpha", sortedKeyedDup(1)(1))
    Call AssertEqual("CollectionBubbleSort keyed Dup - second key", 2, sortedKeyedDup(2)(0))
    Call AssertEqual("CollectionBubbleSort keyed Dup - third key", 2, sortedKeyedDup(3)(0))

    ' ===== ArrayBubbleSort tests - array of array(key, payload) =====
    Dim arrKeyedAsc As Variant
    arrKeyedAsc = Array( _
        Array(3, "Charlie"), _
        Array(1, "Alpha"), _
        Array(2, "Bravo") _
    )

    Dim sortedArrKeyedAsc As Variant
    sortedArrKeyedAsc = ArrayBubbleSort(arrKeyedAsc, True)

    Call AssertEqual("ArrayBubbleSort keyed Asc - first key", 1, sortedArrKeyedAsc(0)(0))
    Call AssertEqual("ArrayBubbleSort keyed Asc - first payload", "Alpha", sortedArrKeyedAsc(0)(1))
    Call AssertEqual("ArrayBubbleSort keyed Asc - last key", 3, sortedArrKeyedAsc(UBound(sortedArrKeyedAsc))(0))
    Call AssertEqual("ArrayBubbleSort keyed Asc - last payload", "Charlie", sortedArrKeyedAsc(UBound(sortedArrKeyedAsc))(1))

    Dim arrKeyedDesc As Variant
    arrKeyedDesc = Array( _
        Array(3, "Charlie"), _
        Array(1, "Alpha"), _
        Array(2, "Bravo") _
    )

    Dim sortedArrKeyedDesc As Variant
    sortedArrKeyedDesc = ArrayBubbleSort(arrKeyedDesc, False)

    Call AssertEqual("ArrayBubbleSort keyed Desc - first key", 3, sortedArrKeyedDesc(0)(0))
    Call AssertEqual("ArrayBubbleSort keyed Desc - first payload", "Charlie", sortedArrKeyedDesc(0)(1))
    Call AssertEqual("ArrayBubbleSort keyed Desc - last key", 1, sortedArrKeyedDesc(UBound(sortedArrKeyedDesc))(0))
    Call AssertEqual("ArrayBubbleSort keyed Desc - last payload", "Alpha", sortedArrKeyedDesc(UBound(sortedArrKeyedDesc))(1))

    ' ===== ArrayBubbleSort tests - single keyed item =====
    Dim arrKeyedSingle As Variant
    arrKeyedSingle = Array(Array(42, "Only"))
    
    Dim sortedArrKeyedSingle As Variant
    sortedArrKeyedSingle = ArrayBubbleSort(arrKeyedSingle, True)
    
    Call AssertEqual("ArrayBubbleSort keyed Single - key", 42, sortedArrKeyedSingle(0)(0))
    Call AssertEqual("ArrayBubbleSort keyed Single - payload", "Only", sortedArrKeyedSingle(0)(1))
End Sub

' Re-written by copilot on 2025-08-15
Function IndexOf(Value As Variant, source As Variant) As Long
    Dim i As Long

    If IsArray(source) Then
        ' Handle array
        On Error GoTo ArrayBoundsError
        For i = LBound(source) To UBound(source)
            If source(i) = Value Then
                IndexOf = i
                Exit Function
            End If
        Next i
        IndexOf = -1
        Exit Function

ArrayBoundsError:
        IndexOf = -1
        Exit Function

    ElseIf TypeName(source) = "Collection" Then
        ' Handle collection (1-based)
        For i = 1 To source.Count
            If source(i) = Value Then
                IndexOf = i
                Exit Function
            End If
        Next i
        IndexOf = -1
        Exit Function

    Else
        ' Unsupported type
        IndexOf = -1
    End If
End Function

' Re-written by copilot on 2025-08-15
' New logic is: convert the Collection to an array, then sort, then back to collection
Function CollectionBubbleSort(colInput As Collection, Optional ascending As Boolean = True) As Collection
    Dim arrTemp() As Variant
    Dim i As Long
    Dim sortedArr As Variant
    Dim colSorted As New Collection

    ' Handle empty collection
    If colInput.Count = 0 Then
        Set CollectionBubbleSort = colSorted
        Exit Function
    End If

    ' Convert collection to array
    ReDim arrTemp(0 To colInput.Count - 1)
    For i = 1 To colInput.Count
        arrTemp(i - 1) = colInput(i)
    Next i

    ' Sort array
    sortedArr = ArrayBubbleSort(arrTemp, ascending)

    ' Convert back to collection
    For i = LBound(sortedArr) To UBound(sortedArr)
        colSorted.Add sortedArr(i)
    Next i

    Set CollectionBubbleSort = colSorted
End Function

' Added by copilot on 2025-08-15
' Extended by ChatGPT 5.3 to handle Arrays of Arrays as well as Arrays of Data
Function ArrayBubbleSort(arrInput As Variant, Optional ascending As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim arrSorted() As Variant
    Dim leftValue As Variant
    Dim rightValue As Variant

    If Not IsArray(arrInput) Or UBound(arrInput) < LBound(arrInput) Then
        ArrayBubbleSort = arrInput
        Exit Function
    End If
    
    arrSorted = arrInput

    For i = LBound(arrSorted) To UBound(arrSorted) - 1
        For j = LBound(arrSorted) To UBound(arrSorted) - 1 - i
            
            leftValue = SortValue(arrSorted(j))
            rightValue = SortValue(arrSorted(j + 1))
            
            If (ascending And leftValue > rightValue) Or _
               (Not ascending And leftValue < rightValue) Then
                
                temp = arrSorted(j)
                arrSorted(j) = arrSorted(j + 1)
                arrSorted(j + 1) = temp
            End If
            
        Next j
    Next i

    ArrayBubbleSort = arrSorted
End Function

Private Function SortValue(ByVal v As Variant) As Variant
    If IsArray(v) Then
        SortValue = v(0)
    Else
        SortValue = v
    End If
End Function

