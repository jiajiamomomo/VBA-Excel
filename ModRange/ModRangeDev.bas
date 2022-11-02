Attribute VB_Name = "ModRangeDev"
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



'
'this group of functions use Worksheet.UsedRange property
'
'sht As Worksheet
'when the entire worksheet is empty in value & formula & format, UsedRange returns range("A1")
'value & formula & format all affect UsedRange
Function LastRow1(sht)
Set r = sht.UsedRange
LastRow1 = r.Rows(r.Rows.Count).row
End Function

Function LastColumn1(sht)
Set r = sht.UsedRange
LastColumn1 = r.Columns(r.Columns.Count).Column
End Function

Function FirstRow1(sht)
Set r = sht.UsedRange
FirstRow1 = r.Rows(1).row
End Function

Function FirstColumn1(sht)
Set r = sht.UsedRange
FirstColumn1 = r.Columns(1).Column
End Function



'
'this group of functions used Range.SpecialCells method
'
'sht As Worksheet
'when the entire worksheet is empty in value & formula & format, SpecialCells returns range("A1")
'value & formula & format all affect SpecialCells
Function LastRow2(sht)
LastRow2 = sht.Cells.SpecialCells(xlCellTypeLastCell).row
End Function

Function LastColumn2(sht)
LastColumn2 = sht.Cells.SpecialCells(xlCellTypeLastCell).Column
End Function




Sub TestLast()
Set sht = ThisWorkbook.Worksheets("test")
sht.Activate
Call TestLast_(sht)

Set sht = ThisWorkbook.Worksheets("empty")
sht.Activate
Call TestLast_(sht)

Set sht = ThisWorkbook.Worksheets("test1")
sht.Activate
Call TestLast_(sht)

Set sht = ThisWorkbook.Worksheets("rand")
sht.Activate
Call TestLast_(sht)
End Sub

'sht as Worksheet
Sub TestLast_(sht)
msg = ""

c = 3
r = LastRowInColumn(sht, c)
If r = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
Else
    msg = msg & "LastRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastCellInColumn(sht, c)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastCellInColumn(sht, " & CStr(c) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastCellInColumn(sht, " & CStr(c) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

c = 3
r = FirstRowInColumn(sht, c)
If r = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
Else
    msg = msg & "FirstRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstCellInColumn(sht, c)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstCellInColumn(sht, " & CStr(c) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstCellInColumn(sht, " & CStr(c) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

c = 5
r = LastRowInColumn(sht, c)
If r = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
Else
    msg = msg & "LastRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastCellInColumn(sht, c)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastCellInColumn(sht, " & CStr(c) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastCellInColumn(sht, " & CStr(c) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

c = 5
r = FirstRowInColumn(sht, c)
If r = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
Else
    msg = msg & "FirstRowInColumn(sht, " & CStr(c) & ") = " & CStr(r) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstCellInColumn(sht, c)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(c) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstCellInColumn(sht, " & CStr(c) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstCellInColumn(sht, " & CStr(c) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

r = 4
c = LastColumnInRow(sht, r)
If c = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(r) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastColumnInRow(sht, " & CStr(r) & ") = " & CStr(c) & Chr(13)
Else
    msg = msg & "LastColumnInRow(sht, " & CStr(r) & ") = " & CStr(c) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastCellInRow(sht, r)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(r) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastCellInRow(sht, " & CStr(r) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastCellInRow(sht, " & CStr(r) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

r = 4
c = FirstColumnInRow(sht, r)
If c = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(r) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstColumnInRow(sht, " & CStr(r) & ") = " & CStr(c) & Chr(13)
Else
    msg = msg & "FirstColumnInRow(sht, " & CStr(r) & ") = " & CStr(c) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstCellInRow(sht, r)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(r) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstCellInRow(sht, " & CStr(r) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstCellInRow(sht, " & CStr(r) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

MsgBox msg
msg = ""

r = LastRow(sht)
If r = 0 Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "LastRow(sht) = " & CStr(r) & Chr(13)
Else
    msg = msg & "LastRow(sht) = " & CStr(r) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastCellInLastRow(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "LastCellInLastRow(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of LastCellInLastRow(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

c = LastColumn(sht)
If c = 0 Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "LastColumn(sht) = " & CStr(c) & Chr(13)
Else
    msg = msg & "LastColumn(sht) = " & CStr(c) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastCellInLastColumn(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "LastCellInLastColumn(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of LastCellInLastColumn(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

r = FirstRow(sht)
If r = 0 Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "FirstRow(sht) = " & CStr(r) & Chr(13)
Else
    msg = msg & "FirstRow(sht) = " & CStr(r) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstCellInFirstRow(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "FirstCellInFirstRow(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstCellInFirstRow(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

c = FirstColumn(sht)
If c = 0 Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "FirstColumn(sht) = " & CStr(c) & Chr(13)
Else
    msg = msg & "FirstColumn(sht) = " & CStr(c) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstCellInFirstColumn(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "FirstCellInFirstColumn(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstCellInFirstColumn(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

MsgBox msg
msg = ""

r = LastRow1(sht)
msg = msg & "LastRow1(sht) = " & CStr(r) & Chr(13)
msg = msg & Chr(13)

c = LastColumn1(sht)
msg = msg & "LastColumn1(sht) = " & CStr(c) & Chr(13)
msg = msg & Chr(13)

r = LastRow2(sht)
msg = msg & "LastRow2(sht) = " & CStr(r) & Chr(13)
msg = msg & Chr(13)

c = LastColumn2(sht)
msg = msg & "LastColumn2(sht) = " & CStr(c) & Chr(13)
msg = msg & Chr(13)

MsgBox msg
End Sub



'In the specified direction, Range.End method finds first or last cell in a continuous range which
'is non-empty in value or formula
'Format won't affect Range.End method
'Range.End method is equivalent to pressing END+UP key, END+DOWN key, END+LEFT key, or END+RIGHT key
'Keep executing Range.End(xlUp) eventually stops at the very first row. It is always row 1
'Keep executing Range.End(xlDown) eventually stops at the very last possible row. It is always
'row Rows.Count. Actual row number varies in different EXCEL version, and it is row 1048567 in EXCEL 365
'Keep executing Range.End(xlToLeft) eventually stops at the very first column. It is always column "A"
'Keep executing Range.End(xlToRight) eventually stops at the very last possible column. It is always
'column Columns.Count. Actual column number varies in different EXCEL version, and it is column "XFD" in EXCEL 365
Sub TestEnd()
Dim sht
Set sht = Worksheets("test")
sht.Activate

sht.Range("C7").End(xlUp).Select
MsgBox "sht.Range('C7').End(xlUp) is selected"
sht.Range("C7").End(xlDown).Select
MsgBox "sht.Range('C7').End(xlDown) is selected"
sht.Range("F4").End(xlToLeft).Select
MsgBox "sht.Range('F4').End(xlToLeft) is selected"
sht.Range("F4").End(xlToRight).Select
MsgBox "sht.Range('F4').End(xlToRight) is selected"

sht.Range("L8").End(xlUp).Select
MsgBox "sht.Range('L8').End(xlUp) is selected"
sht.Range("L8").End(xlDown).Select
MsgBox "sht.Range('L8').End(xlDown) is selected"
End Sub



'
'using SpecialCells to exactly find out all cells with specified cell type & value
'
Sub TestSpecialCells()
Dim sht
Set sht = ThisWorkbook.Worksheets("test")
sht.Activate
Call TestSpecialCells_(sht)

Set sht = ThisWorkbook.Worksheets("empty")
sht.Activate
Call TestSpecialCells_(sht)

Set sht = ThisWorkbook.Worksheets("rand")
sht.Activate
Call TestSpecialCells_(sht)
End Sub

'sht As Worksheet
Sub TestSpecialCells_(sht)
Dim rng

'when the entire worksheet is empty in value & formula & format, UsedRange returns range("A1")
'value & formula & format all affect UsedRange
Set rng = sht.UsedRange
rng.Select
MsgBox "sht.UsedRange"

'XlCellType
'when the entire worksheet is empty in value & formula & format, UsedRange returns range("A1")
'value & formula & format all affect SpecialCells(xlCellTypeLastCell)
Set rng = sht.Cells.SpecialCells(xlCellTypeLastCell)
rng.Select
MsgBox "sht.Cells.SpecialCells(xlCellTypeLastCell)"

'Set rng = Nothing is important and necessary
'if the specified cell type is nod found, then rng will keep its previous value
On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeConstants)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants)"
End If
On Error GoTo 0

On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeFormulas)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas)"
End If
On Error GoTo 0

On Error Resume Next
Set rng = Nothing
Set rng = sht.UsedRange.SpecialCells(xlCellTypeBlanks)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeBlanks) is not found"
Else
    rng.Select
    MsgBox "sht.UsedRange.SpecialCells(xlCellTypeBlanks)"
End If
On Error GoTo 0

On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeConstants, xlLogical)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants, xlLogical) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants, xlLogical)"
End If
On Error GoTo 0

'xlNumbers means data types Integer, Long, Single, Double, Currency, Decimal, Date
On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeConstants, xlNumbers)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants, xlNumbers) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants, xlNumbers)"
End If
On Error GoTo 0

On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeConstants, xlTextValues)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants, xlTextValues) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeConstants, xlTextValues)"
End If
On Error GoTo 0

On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeFormulas, xlLogical)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas, xlLogical) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas, xlLogical)"
End If
On Error GoTo 0

'xlNumbers means data types Integer, Long, Single, Double, Currency, Decimal, Date
On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeFormulas, xlNumbers)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas, xlNumbers) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas, xlNumbers)"
End If
On Error GoTo 0

On Error Resume Next
Set rng = Nothing
Set rng = sht.Cells.SpecialCells(xlCellTypeFormulas, xlTextValues)
If rng Is Nothing Then
    sht.Range("A10").Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas, xlTextValues) is not found"
Else
    rng.Select
    MsgBox "sht.Cells.SpecialCells(xlCellTypeFormulas, xlTextValues)"
End If
On Error GoTo 0

sht.Range("F14").Select
End Sub








'
'this group of functions use Range.Find method to find non-empty cell
'
'sht As Worksheet
'when the entire worksheet is empty in both value and formula, function returns 0
'format won't affect this function
'use following code to handle empty row/column:
'var = FirstNonEmptyRowInColumn(sht, ColumnIndex)
'If var = 0 Then
'    'not found, do something for empty row/column
'Else
'    'found, do something for non-empty row/column
'End If
Function FirstNonEmptyRowInColumn(sht, ColumnIndex)
Dim cell
Set cell = FirstNonEmptyCellInColumn(sht, ColumnIndex)
If cell Is Nothing Then
    FirstNonEmptyRowInColumn = 0
Else
    FirstNonEmptyRowInColumn = cell.row
End If
End Function

Function FirstNonEmptyColumnInRow(sht, RowIndex)
Dim cell
Set cell = FirstNonEmptyCellInRow(sht, RowIndex)
If cell Is Nothing Then
    FirstNonEmptyColumnInRow = 0
Else
    FirstNonEmptyColumnInRow = cell.Column
End If
End Function

Function LastNonEmptyRowInColumn(sht, ColumnIndex)
Dim cell
Set cell = LastNonEmptyCellInColumn(sht, ColumnIndex)
If cell Is Nothing Then
    LastNonEmptyRowInColumn = 0
Else
    LastNonEmptyRowInColumn = cell.row
End If
End Function

Function LastNonEmptyColumnInRow(sht, RowIndex)
Dim cell
Set cell = LastNonEmptyCellInRow(sht, RowIndex)
If cell Is Nothing Then
    LastNonEmptyColumnInRow = 0
Else
    LastNonEmptyColumnInRow = cell.Column
End If
End Function

'sht As Worksheet
'when the entire row/column is empty in both value and formula, function returns Nothing
'format won't affect this function
'use following code to handle empty row/column:
'set var = FirstNonEmptyCellInColumn(sht, ColumnIndex)
'If var Is Nothing Then
'    'not found, do something for empty row/column
'Else
'    'found, do something for non-empty row/column
'End If
Function FirstNonEmptyCellInColumn(sht, ColumnIndex)
On Error Resume Next
Set FirstNonEmptyCellInColumn = sht.Columns(ColumnIndex).Find(What:="*", _
                    After:=sht.Cells(sht.Rows.Count, ColumnIndex), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstNonEmptyCellInRow(sht, RowIndex)
On Error Resume Next
Set FirstNonEmptyCellInRow = sht.Rows(RowIndex).Find(What:="*", _
                    After:=sht.Cells(RowIndex, sht.Columns.Count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastNonEmptyCellInColumn(sht, ColumnIndex)
On Error Resume Next
Set LastNonEmptyCellInColumn = sht.Columns(ColumnIndex).Find(What:="*", _
                    After:=sht.Cells(1, ColumnIndex), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastNonEmptyCellInRow(sht, RowIndex)
On Error Resume Next
Set LastNonEmptyCellInRow = sht.Rows(RowIndex).Find(What:="*", _
                    After:=sht.Cells(RowIndex, 1), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstNonEmptyCellByRows(sht)
On Error Resume Next
Set FirstNonEmptyCellByRows = sht.Cells.Find(What:="*", _
                    After:=sht.Cells(sht.Rows.Count, sht.Columns.Count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstNonEmptyCellByColumns(sht)
On Error Resume Next
Set FirstNonEmptyCellByColumns = sht.Cells.Find(What:="*", _
                    After:=sht.Cells(sht.Rows.Count, sht.Columns.Count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastNonEmptyCellByRows(sht)
On Error Resume Next
Set LastNonEmptyCellByRows = sht.Cells.Find(What:="*", _
                    After:=sht.Cells(1, 1), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastNonEmptyCellByColumns(sht)
On Error Resume Next
Set LastNonEmptyCellByColumns = sht.Cells.Find(What:="*", _
                    After:=sht.Cells(1, 1), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function



'
'this group of functions use Range.Find method to find empty cell
'
'sht As Worksheet
'when the entire worksheet is non-empty in both value and formula, function returns 0
'format won't affect this function
'use following code to handle non-empty row/column:
'var = FirstEmptyRowInColumn(sht, ColumnIndex)
'If var = 0 Then
'    'not found, do something for non-empty row/column
'Else
'    'found, do something for empty row/column
'End If
Function FirstEmptyRowInColumn(sht, ColumnIndex)
Dim cell
Set cell = FirstEmptyCellInColumn(sht, ColumnIndex)
If cell Is Nothing Then
    FirstEmptyRowInColumn = 0
Else
    FirstEmptyRowInColumn = cell.row
End If
End Function

Function FirstEmptyColumnInRow(sht, RowIndex)
Dim cell
Set cell = FirstEmptyCellInRow(sht, RowIndex)
If cell Is Nothing Then
    FirstEmptyColumnInRow = 0
Else
    FirstEmptyColumnInRow = cell.Column
End If
End Function

Function LastEmptyRowInColumn(sht, ColumnIndex)
Dim cell
Set cell = LastEmptyCellInColumn(sht, ColumnIndex)
If cell Is Nothing Then
    LastEmptyRowInColumn = 0
Else
    LastEmptyRowInColumn = cell.row
End If
End Function

Function LastEmptyColumnInRow(sht, RowIndex)
Dim cell
Set cell = LastEmptyCellInRow(sht, RowIndex)
If cell Is Nothing Then
    LastEmptyColumnInRow = 0
Else
    LastEmptyColumnInRow = cell.Column
End If
End Function

'sht As Worksheet
'when the entire row/column is non-empty in both value and formula, function returns Nothing
'format won't affect this function
'use following code to handle non-empty row/column:
'set var = FirstNonEmptyCellInColumn(sht, ColumnIndex)
'If var Is Nothing Then
'    'not found, do something for non-empty row/column
'Else
'    'found, do something for empty row/column
'End If
Function FirstEmptyCellInColumn(sht, ColumnIndex)
On Error Resume Next
Set FirstEmptyCellInColumn = sht.Columns(ColumnIndex).Find(What:="", _
                    After:=sht.Cells(sht.Rows.Count, ColumnIndex), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstEmptyCellInRow(sht, RowIndex)
On Error Resume Next
Set FirstEmptyCellInRow = sht.Rows(RowIndex).Find(What:="", _
                    After:=sht.Cells(RowIndex, sht.Columns.Count), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastEmptyCellInColumn(sht, ColumnIndex)
On Error Resume Next
Set LastEmptyCellInColumn = sht.Columns(ColumnIndex).Find(What:="", _
                    After:=sht.Cells(1, ColumnIndex), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastEmptyCellInRow(sht, RowIndex)
On Error Resume Next
Set LastEmptyCellInRow = sht.Rows(RowIndex).Find(What:="", _
                    After:=sht.Cells(RowIndex, 1), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstEmptyCellByRows(sht)
On Error Resume Next
Set FirstEmptyCellByRows = sht.Cells.Find(What:="", _
                    After:=sht.Cells(sht.Rows.Count, sht.Columns.Count), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function FirstEmptyCellByColumns(sht)
On Error Resume Next
Set FirstEmptyCellByColumns = sht.Cells.Find(What:="", _
                    After:=sht.Cells(sht.Rows.Count, sht.Columns.Count), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastEmptyCellByRows(sht)
On Error Resume Next
Set LastEmptyCellByRows = sht.Cells.Find(What:="", _
                    After:=sht.Cells(1, 1), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function

Function LastEmptyCellByColumns(sht)
On Error Resume Next
Set LastEmptyCellByColumns = sht.Cells.Find(What:="", _
                    After:=sht.Cells(1, 1), _
                    LookAt:=xlWhole, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
On Error GoTo 0
End Function



Sub TestFindNonEmpty()
Dim sht
Set sht = ThisWorkbook.Worksheets("test")
sht.Activate
Call TestFindNonEmpty_(sht)

Set sht = ThisWorkbook.Worksheets("empty")
sht.Activate
Call TestFindNonEmpty_(sht)

Set sht = ThisWorkbook.Worksheets("test1")
sht.Activate
Call TestFindNonEmpty_(sht)

Set sht = ThisWorkbook.Worksheets("rand")
sht.Activate
Call TestFindNonEmpty_(sht)
End Sub

'sht as Worksheet
Sub TestFindNonEmpty_(sht)
Dim msg
msg = ""

Dim col
Dim row
Dim cell

col = 3
row = LastNonEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "LastNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastNonEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastNonEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

col = 3
row = FirstNonEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "FirstNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstNonEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstNonEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

col = 5
row = LastNonEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "LastNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastNonEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastNonEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

col = 5
row = FirstNonEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "FirstNonEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstNonEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstNonEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

row = 4
col = LastNonEmptyColumnInRow(sht, row)
If col = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
Else
    msg = msg & "LastNonEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastNonEmptyCellInRow(sht, row)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyCellInRow(sht, " & CStr(row) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastNonEmptyCellInRow(sht, " & CStr(row) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

row = 4
col = FirstNonEmptyColumnInRow(sht, row)
If col = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
Else
    msg = msg & "FirstNonEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstNonEmptyCellInRow(sht, row)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyCellInRow(sht, " & CStr(row) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstNonEmptyCellInRow(sht, " & CStr(row) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

MsgBox msg
msg = ""

Set cell = LastNonEmptyCellByRows(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyCellByRows(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of LastNonEmptyCellByRows(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastNonEmptyCellByColumns(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "LastNonEmptyCellByColumns(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of LastNonEmptyCellByColumns(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstNonEmptyCellByRows(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyCellByRows(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstNonEmptyCellByRows(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstNonEmptyCellByColumns(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is empty in both value and formula" & Chr(13)
    msg = msg & "FirstNonEmptyCellByColumns(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstNonEmptyCellByColumns(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

MsgBox msg
End Sub



Sub TestFindEmpty()
Dim sht
Set sht = ThisWorkbook.Worksheets("test")
sht.Activate
Call TestFindEmpty_(sht)

'when the entire worksheet is ALL empty and blank, Range.Find method always can't find anything, no matter
'the target is empty cell or non-empty cell
Set sht = ThisWorkbook.Worksheets("empty")
sht.Activate
Call TestFindEmpty_(sht)

Set sht = ThisWorkbook.Worksheets("test1")
sht.Activate
Call TestFindEmpty_(sht)

Set sht = ThisWorkbook.Worksheets("rand")
sht.Activate
Call TestFindEmpty_(sht)
End Sub

'sht as Worksheet
Sub TestFindEmpty_(sht)
Dim msg
msg = ""

Dim col
Dim row
Dim cell

col = 3
row = LastEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "LastEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

col = 3
row = FirstEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "FirstEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

col = 5
row = LastEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "LastEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

col = 5
row = FirstEmptyRowInColumn(sht, col)
If row = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
Else
    msg = msg & "FirstEmptyRowInColumn(sht, " & CStr(col) & ") = " & CStr(row) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstEmptyCellInColumn(sht, col)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Column(" & CStr(col) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyCellInColumn(sht, " & CStr(col) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstEmptyCellInColumn(sht, " & CStr(col) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

row = 4
col = LastEmptyColumnInRow(sht, row)
If col = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
Else
    msg = msg & "LastEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastEmptyCellInRow(sht, row)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyCellInRow(sht, " & CStr(row) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of LastEmptyCellInRow(sht, " & CStr(row) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

row = 4
col = FirstEmptyColumnInRow(sht, row)
If col = 0 Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
Else
    msg = msg & "FirstEmptyColumnInRow(sht, " & CStr(row) & ") = " & CStr(col) & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstEmptyCellInRow(sht, row)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "').Row(" & CStr(row) & ") is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyCellInRow(sht, " & CStr(row) & ") = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstEmptyCellInRow(sht, " & CStr(row) & ") = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

MsgBox msg
msg = ""

Set cell = LastEmptyCellByRows(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyCellByRows(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of LastEmptyCellByRows(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

Set cell = LastEmptyCellByColumns(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is non-empty in both value and formula" & Chr(13)
    msg = msg & "LastEmptyCellByColumns(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of LastEmptyCellByColumns(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstEmptyCellByRows(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyCellByRows(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstEmptyCellByRows(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

Set cell = FirstEmptyCellByColumns(sht)
If cell Is Nothing Then
    msg = msg & "worksheet('" & sht.Name & "') is non-empty in both value and formula" & Chr(13)
    msg = msg & "FirstEmptyCellByColumns(sht) = Nothing" & Chr(13)
Else
    msg = msg & "address of FirstEmptyCellByColumns(sht) = " & cell.Address & Chr(13)
End If
msg = msg & Chr(13)

MsgBox msg
End Sub



Sub FillWorkSheet()
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets("rand")
sht.Activate

'set as sparse array to save running time
Const ratio As Single = 0.00001

Dim RowsCount, ColumnsCount As Long
RowsCount = sht.Rows.Count
ColumnsCount = sht.Columns.Count
Dim row, col As Long

Dim sRowsCount As String
sRowsCount = CStr(RowsCount)

Application.ScreenUpdating = False
Randomize
sht.Cells.Clear
For row = 1 To RowsCount
    Application.StatusBar = CStr(row) & " / " & sRowsCount
    For col = 1 To ColumnsCount
        If Rnd < ratio Then
            sht.Cells(row, col).Value = True
        End If
    Next
Next
Application.ScreenUpdating = True
End Sub
