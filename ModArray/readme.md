
# ModArray
Description: Array operation. Only sorting so far.
License: MIT, Free Software

## 1D Array
*Parameter:*
*vArray: 1D array*

*Description: sort 1D array in ascending order* <br>
Sub ArraySortASC(ByRef vArray)

*Description: sort 1D array in descending order*
Sub ArraySortDESC(ByRef vArray)

## 2D Array
*Parameter:*
*vArray: 2D array*
*IdxSort_dim1: Sort index in 1st dimension*
*IdxSort_dim2: Sort index in 2nd dimension*

*Description: sort 2D array in ascending order, IdxSort_dim1 is in 1st dimension*
Sub Array2DSortASC_dim1(ByRef vArray, ByVal IdxSort_dim1)

*Description: sort 2D array in descending order, IdxSort_dim1 is in 1st dimension*
Sub Array2DSortDESC_dim1(ByRef vArray, ByVal IdxSort_dim1)

*Description: sort 2D array in ascending order, IdxSort_dim2 is in 2nd dimension*
Sub Array2DSortASC_dim2(ByRef vArray, ByVal IdxSort_dim2)

*Description: sort 2D array in descending order, IdxSort_dim2 is in 2nd dimension*
Sub Array2DSortDESC_dim2(ByRef vArray, ByVal IdxSort_dim2
