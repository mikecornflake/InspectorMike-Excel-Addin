Attribute VB_Name = "libArray"
' 2025 08 15 - Added ArrayBubbleSort
'            - This unit handles both Collections and Arrays, they have different indexing rules (array 0 based, collections 1 based).
'            -     IndexOf modified to handle the differences

Option Explicit

Private Const INSERTIONSORT_THRESHOLD As Long = 7

' Can't do this in a procedure, but leaving code here so I don't have to google how to increase array again...
'
'Function Add(ByVal AValue As Variant, ByRef Aarr As Variant) As Long
'    ReDim Preserve Aarr(UBound(Aarr) + 1)
'    Aarr(UBound(Aarr)) = AValue
'    Add = UBound(Aarr)
'End Function

' Re-written by copilot on 2025-08-15
Public Function Array_IndexOf(ByRef pArray As Variant, ByVal pValue As Variant) As Long
    Dim i As Long
    
    On Error GoTo ArrayBoundsError
    For i = LBound(pArray) To UBound(pArray)
        If pArray(i) = pValue Then
            Array_IndexOf = i
            Exit Function
        End If
    Next i
    
    Array_IndexOf = -1
    
    Exit Function

ArrayBoundsError:
    Array_IndexOf = -1
End Function

Public Function Collection_IndexOf(ByVal pCollection As Collection, ByVal pValue As Variant) As Long
    Dim i As Long
    
    If TypeName(pCollection) = "Collection" Then
        ' Handle collection (1-based)
        For i = 1 To pCollection.Count
            If pCollection(i) = pValue Then
                Collection_IndexOf = i
                Exit Function
            End If
        Next i
        
        Collection_IndexOf = -1
        Exit Function
    Else
        ' Unsupported type
        Collection_IndexOf = -1
    End If
End Function

' Re-written by copilot on 2025-08-15
' New logic is: convert the Collection to an array, then sort, then back to collection
Public Function Collection_Sort(ByVal pCollection As Collection, Optional ByVal pAscending As Boolean = True) As Collection
    Dim arrTemp() As Variant
    Dim i As Long
    Dim sortedArr As Variant
    Dim colSorted As New Collection

    ' Handle empty collection
    If pCollection.Count = 0 Then
        Set Collection_Sort = colSorted
        Exit Function
    End If

    ' Convert collection to array
    ReDim arrTemp(0 To pCollection.Count - 1)
    For i = 1 To pCollection.Count
        arrTemp(i - 1) = pCollection(i)
    Next i

    ' Sort array
    sortedArr = Array_Sort(arrTemp, pAscending)

    ' Convert back to collection
    For i = LBound(sortedArr) To UBound(sortedArr)
        colSorted.Add sortedArr(i)
    Next i

    Set Collection_Sort = colSorted
End Function

' Added by copilot on 2025-08-15
' Extended by ChatGPT 5.3 to handle Arrays of Arrays as well as Arrays of Data
Public Function Array_Sort(ByRef pArray As Variant, Optional ByVal pAscending As Boolean = True) As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim arrSorted() As Variant
    Dim leftValue As Variant
    Dim rightValue As Variant

    If Not IsArray(pArray) Or UBound(pArray) < LBound(pArray) Then
        Array_Sort = pArray
        Exit Function
    End If
    
    arrSorted = pArray

    For i = LBound(arrSorted) To UBound(arrSorted) - 1
        For j = LBound(arrSorted) To UBound(arrSorted) - 1 - i
            
            leftValue = SortValue(arrSorted(j))
            rightValue = SortValue(arrSorted(j + 1))
            
            If (pAscending And leftValue > rightValue) Or _
               (Not pAscending And leftValue < rightValue) Then
                
                temp = arrSorted(j)
                arrSorted(j) = arrSorted(j + 1)
                arrSorted(j + 1) = temp
            End If
            
        Next j
    Next i

    Array_Sort = arrSorted
End Function

Private Function SortValue(ByVal v As Variant) As Variant
    If IsArray(v) Then
        SortValue = v(0)
    Else
        SortValue = v
    End If
End Function

