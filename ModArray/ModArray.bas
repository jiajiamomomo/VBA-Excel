Option Explicit

'sort 1D array in ascending order
Sub ArraySortASC(ByRef vArray, Optional inLow, Optional inHi)
If IsMissing(inLow) Or IsMissing(inHi) Then
    inLow = LBound(vArray)
    inHi = UBound(vArray)
End If

Dim tmpLow As Long
Dim tmpHi As Long
tmpLow = inLow
tmpHi = inHi

Dim pivot
pivot = vArray((inLow + inHi) \ 2)

While (tmpLow <= tmpHi)
    While (vArray(tmpLow) < pivot And _
           tmpLow < inHi)
        tmpLow = tmpLow + 1
    Wend
    
    While (pivot < vArray(tmpHi) And _
           inLow < tmpHi)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        Call SwapArrayElement(vArray, tmpLow, tmpHi)
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Wend

If (inLow < tmpHi) Then
    Call ArraySortASC(vArray, inLow, tmpHi)
End If

If (tmpLow < inHi) Then
    Call ArraySortASC(vArray, tmpLow, inHi)
End If
End Sub

'sort 1D array in descending order
Sub ArraySortDESC(ByRef vArray, Optional inLow, Optional inHi)
If IsMissing(inLow) Or IsMissing(inHi) Then
    inLow = LBound(vArray)
    inHi = UBound(vArray)
End If

Dim tmpLow As Long
Dim tmpHi As Long
tmpLow = inLow
tmpHi = inHi

Dim pivot
pivot = vArray((inLow + inHi) \ 2)

While (tmpLow <= tmpHi)
    While (vArray(tmpLow) > pivot And _
           tmpLow < inHi)
        tmpLow = tmpLow + 1
    Wend
    
    While (pivot > vArray(tmpHi) And _
           inLow < tmpHi)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        Call SwapArrayElement(vArray, tmpLow, tmpHi)
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Wend

If (inLow < tmpHi) Then
    Call ArraySortDESC(vArray, inLow, tmpHi)
End If

If (tmpLow < inHi) Then
    Call ArraySortDESC(vArray, tmpLow, inHi)
End If
End Sub

'swap element in 1D array
Private Sub SwapArrayElement(ByRef vArray, ByVal Idx1 As Long, ByVal Idx2 As Long)
Dim tmp
tmp = vArray(Idx1)
vArray(Idx1) = vArray(Idx2)
vArray(Idx2) = tmp
End Sub

Private Sub test()
Dim a
Dim i

a = [{9,1,8,2,7,3,6,4,5}]
Call ArraySortASC(a)

For i = LBound(a) To UBound(a)
    Debug.Print i, a(i)
Next
Debug.Print

a = [{9,1,8,2,7,3,6,4,5}]
Call ArraySortDESC(a)

For i = LBound(a) To UBound(a)
    Debug.Print i, a(i)
Next
Debug.Print
End Sub





'sort 2D array in ascending order, IdxSort_dim1 is in 1st dimension
Sub Array2DSortASC_dim1(ByRef vArray, ByVal IdxSort_dim1, Optional inLow, Optional inHi)
If IsMissing(inLow) Or IsMissing(inHi) Then
    inLow = LBound(vArray, 2)
    inHi = UBound(vArray, 2)
End If

Dim tmpLow As Long
Dim tmpHi As Long
tmpLow = inLow
tmpHi = inHi

Dim pivot
pivot = vArray(IdxSort_dim1, (inLow + inHi) \ 2)

While (tmpLow <= tmpHi)
    While (vArray(IdxSort_dim1, tmpLow) < pivot And _
           tmpLow < inHi)
        tmpLow = tmpLow + 1
    Wend
    
    While (pivot < vArray(IdxSort_dim1, tmpHi) And _
           inLow < tmpHi)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        Call SwapArray2DElement_dim1(vArray, tmpLow, tmpHi)
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Wend

If (inLow < tmpHi) Then
    Call Array2DSortASC_dim1(vArray, IdxSort_dim1, inLow, tmpHi)
End If

If (tmpLow < inHi) Then
    Call Array2DSortASC_dim1(vArray, IdxSort_dim1, tmpLow, inHi)
End If
End Sub

'sort 2D array in descending order, IdxSort_dim1 is in 1st dimension
Sub Array2DSortDESC_dim1(ByRef vArray, ByVal IdxSort_dim1, Optional inLow, Optional inHi)
If IsMissing(inLow) Or IsMissing(inHi) Then
    inLow = LBound(vArray, 2)
    inHi = UBound(vArray, 2)
End If

Dim tmpLow As Long
Dim tmpHi As Long
tmpLow = inLow
tmpHi = inHi

Dim pivot
pivot = vArray(IdxSort_dim1, (inLow + inHi) \ 2)

While (tmpLow <= tmpHi)
    While (vArray(IdxSort_dim1, tmpLow) > pivot And _
           tmpLow < inHi)
        tmpLow = tmpLow + 1
    Wend
    
    While (pivot > vArray(IdxSort_dim1, tmpHi) And _
           inLow < tmpHi)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        Call SwapArray2DElement_dim1(vArray, tmpLow, tmpHi)
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Wend

If (inLow < tmpHi) Then
    Call Array2DSortDESC_dim1(vArray, IdxSort_dim1, inLow, tmpHi)
End If

If (tmpLow < inHi) Then
    Call Array2DSortDESC_dim1(vArray, IdxSort_dim1, tmpLow, inHi)
End If
End Sub

'swap element in 2D array, Idx1 & Idx2 is in 2nd dimension
Private Sub SwapArray2DElement_dim1(ByRef vArray, ByVal Idx1 As Long, ByVal Idx2 As Long)
Dim i
For i = LBound(vArray, 1) To UBound(vArray, 1)
    Dim tmp
    tmp = vArray(i, Idx1)
    vArray(i, Idx1) = vArray(i, Idx2)
    vArray(i, Idx2) = tmp
Next
End Sub





'sort 2D array in ascending order, IdxSort_dim2 is in 2nd dimension
Sub Array2DSortASC_dim2(ByRef vArray, ByVal IdxSort_dim2, Optional inLow, Optional inHi)
If IsMissing(inLow) Or IsMissing(inHi) Then
    inLow = LBound(vArray, 2)
    inHi = UBound(vArray, 2)
End If

Dim tmpLow As Long
Dim tmpHi As Long
tmpLow = inLow
tmpHi = inHi

Dim pivot
pivot = vArray((inLow + inHi) \ 2, IdxSort_dim2)

While (tmpLow <= tmpHi)
    While (vArray(tmpLow, IdxSort_dim2) < pivot And _
           tmpLow < inHi)
        tmpLow = tmpLow + 1
    Wend
    
    While (pivot < vArray(tmpHi, IdxSort_dim2) And _
           inLow < tmpHi)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        Call SwapArray2DElement_dim2(vArray, tmpLow, tmpHi)
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Wend

If (inLow < tmpHi) Then
    Call Array2DSortASC_dim2(vArray, IdxSort_dim2, inLow, tmpHi)
End If

If (tmpLow < inHi) Then
    Call Array2DSortASC_dim2(vArray, IdxSort_dim2, tmpLow, inHi)
End If
End Sub

'sort 2D array in descending order, IdxSort_dim2 is in 2nd dimension
Sub Array2DSortDESC_dim2(ByRef vArray, ByVal IdxSort_dim2, Optional inLow, Optional inHi)
If IsMissing(inLow) Or IsMissing(inHi) Then
    inLow = LBound(vArray, 2)
    inHi = UBound(vArray, 2)
End If

Dim tmpLow As Long
Dim tmpHi As Long
tmpLow = inLow
tmpHi = inHi

Dim pivot
pivot = vArray((inLow + inHi) \ 2, IdxSort_dim2)

While (tmpLow <= tmpHi)
    While (vArray(tmpLow, IdxSort_dim2) > pivot And _
           tmpLow < inHi)
        tmpLow = tmpLow + 1
    Wend
    
    While (pivot > vArray(tmpHi, IdxSort_dim2) And _
           inLow < tmpHi)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        Call SwapArray2DElement_dim2(vArray, tmpLow, tmpHi)
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If
Wend

If (inLow < tmpHi) Then
    Call Array2DSortDESC_dim2(vArray, IdxSort_dim2, inLow, tmpHi)
End If

If (tmpLow < inHi) Then
    Call Array2DSortDESC_dim2(vArray, IdxSort_dim2, tmpLow, inHi)
End If
End Sub

'swap element in 2D array, Idx1 & Idx2 is in 1st dimension
Private Sub SwapArray2DElement_dim2(ByRef vArray, ByVal Idx1 As Long, ByVal Idx2 As Long)
Dim i
For i = LBound(vArray, 2) To UBound(vArray, 2)
    Dim tmp
    tmp = vArray(Idx1, i)
    vArray(Idx1, i) = vArray(Idx2, i)
    vArray(Idx2, i) = tmp
Next
End Sub





Private Sub test2D()
Dim a
Dim i, j



Debug.Print "Befor sort:"
a = [{7, 5, 7, 8, 9; 7, 3, 1, 9, 5; 8, 7, 4, 6, 4; 3, 1, 9, 9, 5; 6, 2, 4, 9, 6}]
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print

Debug.Print "After Array2DSortASC_dim1:"
Call Array2DSortASC_dim1(a, 3)
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print



Debug.Print "Befor sort:"
a = [{7, 5, 7, 8, 9; 7, 3, 1, 9, 5; 8, 7, 4, 6, 4; 3, 1, 9, 9, 5; 6, 2, 4, 9, 6}]
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print

Debug.Print "After Array2DSortDESC_dim1:"
Call Array2DSortDESC_dim1(a, 3)
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print



Debug.Print "Befor sort:"
a = [{7, 5, 7, 8, 9; 7, 3, 1, 9, 5; 8, 7, 4, 6, 4; 3, 1, 9, 9, 5; 6, 2, 4, 9, 6}]
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print

Debug.Print "After Array2DSortASC_dim2:"
Call Array2DSortASC_dim2(a, 3)
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print



Debug.Print "Befor sort:"
a = [{7, 5, 7, 8, 9; 7, 3, 1, 9, 5; 8, 7, 4, 6, 4; 3, 1, 9, 9, 5; 6, 2, 4, 9, 6}]
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print

Debug.Print "After Array2DSortDESC_dim2:"
Call Array2DSortDESC_dim2(a, 3)
For i = LBound(a, 1) To UBound(a, 1)
    Debug.Print i, a(i, 1), a(i, 2), a(i, 3), a(i, 4), a(i, 5)
Next
Debug.Print
End Sub
