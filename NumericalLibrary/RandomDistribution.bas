Attribute VB_Name = "RandomDistribution"
Option Explicit
'
'
Public Sub TestRandom()
Dim sh As Worksheet
Dim r As Range
Dim TimerStart As Double, TimerEnd As Double
Dim maxvalue As Double, minvalue As Double
Dim tmp1 As Double, tmp2 As Double
Dim TmpArray() As Double
Dim i As Integer
Const MaxRow As Integer = 10001
Const n As Integer = 100

Application.ScreenUpdating = False
Set sh = ThisWorkbook.Worksheets("Random")
sh.Columns(1).Clear
sh.Columns(2).Clear
'initiate random generator before calling
'subroutine GaussPolar and GaussRatio
Randomize timer
ReDim TmpArray(2 To MaxRow, 1 To 1)

'the same efficient
Set r = sh.Range(sh.Cells(2, 1), sh.Cells(MaxRow, 1))
sh.Cells(1, 1).Value = "GaussPolar"
TimerStart = timer
For i = 2 To MaxRow
    TmpArray(i, 1) = GaussPolar(0.5)
Next i
r.Value = TmpArray
TimerEnd = timer
MsgBox "GaussPolar " & (TimerEnd - TimerStart) & "secs"

'the same efficient
Set r = sh.Range(sh.Cells(2, 2), sh.Cells(MaxRow, 2))
sh.Cells(1, 2).Value = "GaussRatio"
TimerStart = timer
For i = 2 To MaxRow
    TmpArray(i, 1) = GaussRatio(0.5)
Next i
r.Value = TmpArray
TimerEnd = timer
MsgBox "GaussRatio " & (TimerEnd - TimerStart) & "secs"

sh.Range(sh.Cells(6, 5), sh.Cells(65535, 8)).Clear
For i = 1 To n
    sh.Cells(i + 5, 5).Value = i
Next i

Set r = sh.Range(sh.Cells(2, 1), sh.Cells(MaxRow, 1))
maxvalue = Application.WorksheetFunction.Max(r)
minvalue = Application.WorksheetFunction.Min(r)
tmp1 = 0#
For i = 1 To n
    tmp2 = Application.WorksheetFunction. _
        CountIf(r, _
        "<=" & CStr(minvalue + (maxvalue - minvalue) * i / n))
    sh.Cells(i + 5, 6).Value = tmp2 - tmp1
    tmp1 = tmp2
Next i

Set r = sh.Range(sh.Cells(2, 2), sh.Cells(MaxRow, 2))
maxvalue = Application.WorksheetFunction.Max(r)
minvalue = Application.WorksheetFunction.Min(r)
tmp1 = 0#
For i = 1 To n
    tmp2 = Application.WorksheetFunction. _
        CountIf(r, _
        "<=" & CStr(minvalue + (maxvalue - minvalue) * i / n))
    sh.Cells(i + 5, 7).Value = tmp2 - tmp1
    tmp1 = tmp2
Next i

Set r = sh.Range(sh.Cells(2, 3), sh.Cells(MaxRow, 3))
maxvalue = Application.WorksheetFunction.Max(r)
minvalue = Application.WorksheetFunction.Min(r)
tmp1 = 0#
For i = 1 To n
    tmp2 = Application.WorksheetFunction. _
        CountIf(r, _
        "<=" & CStr(minvalue + (maxvalue - minvalue) * i / n))
    sh.Cells(i + 5, 8).Value = tmp2 - tmp1
    tmp1 = tmp2
Next i

Application.ScreenUpdating = True
End Sub
'Gauss distribution(normal distribution)
'Polar (Box-Mueller) method; See Knuth v2, 3rd ed, p122
Public Function GaussPolar(sigma As Double)
Dim x As Double, y As Double, r2 As Double

Do
'choose x,y in uniform square (-1,-1) to (+1,+1)
    x = -1# + 2# * Rnd()
    y = -1# + 2# * Rnd()
'see if it is in the unit circle
    r2 = x * x + y * y
Loop While (r2 >= 1#)

'Box-Muller transform
GaussPolar = sigma * y * Sqr(-2# * Log(r2) / r2)
End Function
'Gauss distribution(normal distribution)
'Ratio method (Kinderman-Monahan); see Knuth v2, 3rd ed, p130
'K+M, ACM Trans Math Software 3 (1977) 257-260.
Public Function GaussRatio(sigma As Double)
Dim u As Double, v As Double, x As Double

Do
    v = Rnd()
    Do
        u = Rnd()
    Loop While (u = 0)
'Const 1.71552776992141359295 = sqrt(8/e)
'    x = 1.71552776992141359295 * (v - 0.5) / u
    x = 1.71552776992141 * (v - 0.5) / u
Loop While (x * x > -4# * Log(u))

GaussRatio = sigma * x
End Function
