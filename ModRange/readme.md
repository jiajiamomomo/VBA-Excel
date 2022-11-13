
# ModRange
Description: Excel Range. Only first/last row/column/cell so far.<br>
License: MIT, Free Software<br>

## First/Last Row/Column in specified Column/Row
*Parameter:*<br>
*sht: specified worksheet*<br>
*ColumnIndex: specified column index*<br>
*RowIndex: specified row index*<br>

*this group of functions used Range.End method*<br>
*when the entire column/row is empty in both value and formula, function returns 0*<br>
*format won't affect this function*<br>

Function LastRowInColumn(sht, ColumnIndex)<br>

Function LastColumnInRow(sht, RowIndex)<br>

Function FirstRowInColumn(sht, ColumnIndex)<br>

Function FirstColumnInRow(sht, RowIndex)<br>

## First/Last Cell in specified Column/Row
*Parameter:*<br>
*sht: specified worksheet*<br>
*ColumnIndex: specified column index*<br>
*RowIndex: specified row index*<br>

*this group of functions used Range.End method*<br>
*when the entire column/row is empty in both value and formula, function returns 0*<br>
*format won't affect this function*<br>

*use following code to handle empty column/row:*<br>
```vba
set var = LastCellInColumn(sht)
If var Is Nothing Then
    'not found, do something for empty column/row
Else
    'found, do something for non-empty column/row
End If
```
*use IsEmpty(rng) to check if Range rng is empty in both value and formula*<br>

Function LastCellInColumn(sht, ColumnIndex)<br>

Function LastCellInRow(sht, RowIndex)<br>

Function FirstCellInColumn(sht, ColumnIndex)<br>

Function FirstCellInRow(sht, RowIndex)<br>

## First/Last Row/Column in specified Worksheet
*Parameter:*<br>
*sht: specified worksheet*<br>

*this group of functions use Range.Find method*<br>
*when the entire worksheet is empty in both value and formula, function returns 0*<br>
*format won't affect this function*<br>

*use following code to handle empty worksheet:*<br>
```vba
var = LastRow(sht)
If var = 0 Then
    'not found, do something for empty worksheet
Else
    'found, do something for non-empty worksheet
End If
```
Function LastRow(sht)<br>

Function LastColumn(sht)<br>

Function FirstRow(sht)<br>

Function FirstColumn(sht)<br>

## First/Last Cell in specified Worksheet
*Parameter:*<br>
*sht: specified worksheet*<br>

*this group of functions use Range.Find method*<br>
*when the entire worksheet is empty in both value and formula, function returns Nothing*<br>
*format won't affect this function*<br>

*use following code to handle empty worksheet:*<br>
```vba
set var = LastCellInLastRow(sht)
If var Is Nothing Then
    'not found, do something for empty worksheet
Else
    'found, do something for non-empty worksheet
End If
```
Function LastCellInLastRow(sht)<br>

Function LastCellInLastColumn(sht)<br>

Function FirstCellInFirstRow(sht)<br>

Function FirstCellInFirstColumn(sht)<br>
