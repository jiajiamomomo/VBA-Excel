Attribute VB_Name = "ModRange"
Option Explicit
'
'Range Functions & Subs
'
'when we talk about non-empty range, it means value, formula and format in range
'a range may be empty in value, or empty in formula, or empty in format
'a range may be non-empty only in value, or non-empty only in formula, or non-empty only in format



'
'this group of functions used Range.End method
'
'sht As Worksheet
'when the entire column/row is empty in both value and formula, function returns 0
'format won't affect this function
Function LastRowInColumn(sht, ColumnIndex)
Set cell = LastCellInColumn(sht, ColumnIndex)
If cell Is Nothing Then
    LastRowInColumn = 0
Else
    LastRowInColumn = cell.row
End If
End Function

Function LastColumnInRow(sht, RowIndex)
Set cell = LastCellInRow(sht, RowIndex)
If cell Is Nothing Then
    LastColumnInRow = 0
Else
    LastColumnInRow = cell.Column
End If
End Function

Function FirstRowInColumn(sht, ColumnIndex)
Set cell = FirstCellInColumn(sht, ColumnIndex)
If cell Is Nothing Then
    FirstRowInColumn = 0
Else
    FirstRowInColumn = cell.row
End If
End Function

Function FirstColumnInRow(sht, RowIndex)
Set cell = FirstCellInRow(sht, RowIndex)
If cell Is Nothing Then
    FirstColumnInRow = 0
Else
    FirstColumnInRow = cell.Column
End If
End Function

'sht As Worksheet
'when the entire column/row is empty in both value and formula, function returns Nothing
'format won't affect this function
'
'use following code to handle empty column/row:
'set var = LastCellInColumn(sht)
'If var Is Nothing Then
'    'not found, do something for empty worksheet
'Else
'    'found, do something for non-empty worksheet
'End If
'
'use IsEmpty(rng) to check if rng is empty in both value and formula
Function LastCellInColumn(sht, ColumnIndex)
Set cell = sht.Cells(sht.Rows.Count, ColumnIndex)
If IsEmpty(cell) Then
    Set LastCellInColumn = cell.End(xlUp)
    If IsEmpty(LastCellInColumn) Then
        Set LastCellInColumn = Nothing
    End If
Else
    Set LastCellInColumn = cell
End If
End Function

Function LastCellInRow(sht, RowIndex)
Set cell = sht.Cells(RowIndex, sht.Columns.Count)
If IsEmpty(cell) Then
    Set LastCellInRow = cell.End(xlToLeft)
    If IsEmpty(LastCellInRow) Then
        Set LastCellInRow = Nothing
    End If
Else
    Set LastCellInRow = cell
End If
End Function

Function FirstCellInColumn(sht, ColumnIndex)
Set cell = sht.Cells(1, ColumnIndex)
If IsEmpty(cell) Then
    Set FirstCellInColumn = cell.End(xlDown)
    If IsEmpty(FirstCellInColumn) Then
        Set FirstCellInColumn = Nothing
    End If
Else
    Set FirstCellInColumn = cell
End If
End Function

Function FirstCellInRow(sht, RowIndex)
Set cell = sht.Cells(RowIndex, 1)
If IsEmpty(cell) Then
    Set FirstCellInRow = cell.End(xlToRight)
    If IsEmpty(FirstCellInRow) Then
        Set FirstCellInRow = Nothing
    End If
Else
    Set FirstCellInRow = cell
End If
End Function



'
'this group of functions use Range.Find method
'
'sht As Worksheet
'when the entire worksheet is empty in both value and formula, function returns 0
'format won't affect this function
'use following code to handle empty worksheet:
'var = LastRow(sht)
'If var = 0 Then
'    'not found, do something for empty worksheet
'Else
'    'found, do something for non-empty worksheet
'End If
Function LastRow(sht)
Set cell = LastCellInLastRow(sht)
If cell Is Nothing Then
    LastRow = 0
Else
    LastRow = cell.row
End If
End Function

Function LastColumn(sht)
Set cell = LastCellInLastColumn(sht)
If cell Is Nothing Then
    LastColumn = 0
Else
    LastColumn = cell.Column
End If
End Function

Function FirstRow(sht)
Set cell = FirstCellInFirstRow(sht)
If cell Is Nothing Then
    FirstRow = 0
Else
    FirstRow = cell.row
End If
End Function

Function FirstColumn(sht)
Set cell = FirstCellInFirstColumn(sht)
If cell Is Nothing Then
    FirstColumn = 0
Else
    FirstColumn = cell.Column
End If
End Function

'sht As Worksheet
'when the entire worksheet is empty in both value and formula, function returns Nothing
'format won't affect this function
'use following code to handle empty worksheet:
'set var = LastCellInLastRow(sht)
'If var Is Nothing Then
'    'not found, do something for empty worksheet
'Else
'    'found, do something for non-empty worksheet
'End If
Function LastCellInLastRow(sht)
On Error Resume Next
Set LastCellInLastRow = sht.Cells.Find(What:="*", _
                    After:=sht.Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastCellInLastColumn(sht)
On Error Resume Next
Set LastCellInLastColumn = sht.Cells.Find(What:="*", _
                    After:=sht.Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstCellInFirstRow(sht)
On Error Resume Next
Set FirstCellInFirstRow = sht.Cells.Find(What:="*", _
                    After:=sht.Cells(sht.Rows.Count, sht.Columns.Count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstCellInFirstColumn(sht)
On Error Resume Next
Set FirstCellInFirstColumn = sht.Cells.Find(What:="*", _
                    After:=sht.Cells(sht.Rows.Count, sht.Columns.Count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function
