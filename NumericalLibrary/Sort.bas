Attribute VB_Name = "Sort"
'Subroutines with postfix name "Permutation" stores the
'index for sorted array in array "Permutation".
'
'Stable sorting algorithms maintain the relative order
'of records with equal keys (i.e. values). That is, a
'sorting algorithm is stable if whenever there are two
'records R and S with the same key and with R appearing
'before S in the original list, R will appear before S
'in the sorted list.
'Stable:
'Bubble, Insertion, Merge
'Unstable:
'Selection, Shell, Heap, Quick
'The best-performance, stable sorting algorithms is
'Merge.
'The best-performance, unstable sorting algorithms is
'Quick.
Option Explicit

Public Sub TestSort()
Dim n As Long
Dim i As Long, j As Long
Dim jStart As Long, jEnd As Long
Dim r As Long, c As Long
Dim repeats As Long
Const xLM1 As Long = -3, pLM1 As Long = -5
Dim TimerStart As Double, TimerEnd As Double
Dim TimerBubble As Double
Dim TimerSelection As Double
Dim TimerInsertion As Double
Dim TimerShell As Double
Dim TimerHeap As Double
Dim TimerMerge As Double
Dim TimerQuick As Double
Dim x() As Double
Dim Permutation() As Long
Dim message As String
Dim sh As Worksheet
Dim NoTiming As Boolean

'initiate random generator before calling
'subroutine QuickSort and QuickPermutation
Randomize timer

'set NoTiming = False for timing.
'set NoTiming = True to show sort results.
NoTiming = True
If NoTiming Then GoTo LabelOthers

'for timing only
'this takes a long time
Application.ScreenUpdating = False
Set sh = ThisWorkbook.Worksheets("Sort")
For r = 2 To 9
    n = sh.Cells(r, 1).Value
    jEnd = n / 2
    jStart = -(n - (jEnd + 1))
    ReDim x(jStart To jEnd)
    repeats = sh.Cells(r, 9)
    c = 1
    'BubbleSort
    c = c + 1
    Application.StatusBar = "Row: " & r & "   " & "Column: " & c
    TimerBubble = 0#
    For i = 1 To repeats
        For j = jStart To jEnd
            x(j) = Rnd
        Next j
        TimerStart = timer
        Call BubbleSort(x())
        TimerEnd = timer
        TimerBubble = TimerBubble + (TimerEnd - TimerStart)
    Next i
    TimerBubble = TimerBubble / repeats
    sh.Cells(r, c).Value = TimerBubble
    'SelectionSort
    c = c + 1
    Application.StatusBar = "Row: " & r & "   " & "Column: " & c
    TimerSelection = 0#
    For i = 1 To repeats
        For j = jStart To jEnd
            x(j) = Rnd
        Next j
        TimerStart = timer
        Call SelectionSort(x())
        TimerEnd = timer
        TimerSelection = TimerSelection + (TimerEnd - TimerStart)
    Next i
    TimerSelection = TimerSelection / repeats
    sh.Cells(r, c).Value = TimerSelection
    'InsertionSort
    c = c + 1
    Application.StatusBar = "Row: " & r & "   " & "Column: " & c
    TimerInsertion = 0#
    For i = 1 To repeats
        For j = jStart To jEnd
            x(j) = Rnd
        Next j
        TimerStart = timer
        Call InsertionSort(x())
        TimerEnd = timer
        TimerInsertion = TimerInsertion + (TimerEnd - TimerStart)
    Next i
    TimerInsertion = TimerInsertion / repeats
    sh.Cells(r, c).Value = TimerInsertion
    'ShellSort
    c = c + 1
    Application.StatusBar = "Row: " & r & "   " & "Column: " & c
    TimerShell = 0#
    For i = 1 To repeats
        For j = jStart To jEnd
            x(j) = Rnd
        Next j
        TimerStart = timer
        Call ShellSort(x())
        TimerEnd = timer
        TimerShell = TimerShell + (TimerEnd - TimerStart)
    Next i
    TimerShell = TimerShell / repeats
    sh.Cells(r, c).Value = TimerShell
    'HeapSort
    c = c + 1
    Application.StatusBar = "Row: " & r & "   " & "Column: " & c
    TimerHeap = 0#
    For i = 1 To repeats
        For j = jStart To jEnd
            x(j) = Rnd
        Next j
        TimerStart = timer
        Call HeapSort(x())
        TimerEnd = timer
        TimerHeap = TimerHeap + (TimerEnd - TimerStart)
    Next i
    TimerHeap = TimerHeap / repeats
    sh.Cells(r, c).Value = TimerHeap
    'MergeSort
    c = c + 1
    Application.StatusBar = "Row: " & r & "   " & "Column: " & c
    TimerMerge = 0#
    For i = 1 To repeats
        For j = jStart To jEnd
            x(j) = Rnd
        Next j
        TimerStart = timer
        Call MergeSort(x())
        TimerEnd = timer
        TimerMerge = TimerMerge + (TimerEnd - TimerStart)
    Next i
    TimerMerge = TimerMerge / repeats
    sh.Cells(r, c).Value = TimerMerge
    'QuickSort
    c = c + 1
    Application.StatusBar = "Row: " & r & "   " & "Column: " & c
    TimerQuick = 0#
    For i = 1 To repeats
        For j = jStart To jEnd
            x(j) = Rnd
        Next j
        TimerStart = timer
        Call QuickSort(x())
        TimerEnd = timer
        TimerQuick = TimerQuick + (TimerEnd - TimerStart)
    Next i
    TimerQuick = TimerQuick / repeats
    sh.Cells(r, c).Value = TimerQuick
Next r
Application.StatusBar = ""
Application.StatusBar = False
Application.ScreenUpdating = True

Exit Sub

LabelOthers:
'only show sort results

n = 3
ReDim x(-n To n)
jStart = -n
jEnd = n

'BubbleSort
For j = jStart To jEnd
    x(j) = Int((2 * n + 1) * Rnd)
Next j
Call BubbleSort(x())
MsgBox "BubbleSort: " & vbNewLine & ShowPoly(x)
'SelectionSort
For j = jStart To jEnd
    x(j) = Int((2 * n + 1) * Rnd)
Next j
Call SelectionSort(x())
MsgBox "SelectionSort: " & vbNewLine & ShowPoly(x)
'InsertionSort
For j = jStart To jEnd
    x(j) = Int((2 * n + 1) * Rnd)
Next j
Call InsertionSort(x())
MsgBox "InsertionSort: " & vbNewLine & ShowPoly(x)
'ShellSort
For j = jStart To jEnd
    x(j) = Int((2 * n + 1) * Rnd)
Next j
Call ShellSort(x())
MsgBox "ShellSort: " & vbNewLine & ShowPoly(x)
'HeapSort
For j = jStart To jEnd
    x(j) = Int((2 * n + 1) * Rnd)
Next j
Call HeapSort(x())
MsgBox "HeapSort: " & vbNewLine & ShowPoly(x)
'MergeSort
For j = jStart To jEnd
    x(j) = Int((2 * n + 1) * Rnd)
Next j
Call MergeSort(x())
MsgBox "MergeSort: " & vbNewLine & ShowPoly(x)
'QuickSort
For j = jStart To jEnd
    x(j) = Int((2 * n + 1) * Rnd)
Next j
Call QuickSort(x())
MsgBox "QuickSort: " & vbNewLine & ShowPoly(x)

n = 10
ReDim x(1 + xLM1 To n + xLM1)
ReDim Permutation(1 + pLM1 To n + pLM1)
For j = 1 To n
    x(j + xLM1) = Int((2 * n + 1) * Rnd)
Next j

'BubblePermutation
Call BubblePermutation(x(), Permutation())
message = "BubblePermutation x() :" & vbNewLine & _
        ShowPoly(x) & vbNewLine & vbNewLine
For i = 1 To n
    message = message & i & vbTab & Permutation(i + pLM1) & vbTab & x(Permutation(i + pLM1)) & vbNewLine
Next i
MsgBox message
'SelectionPermutation
Call SelectionPermutation(x(), Permutation())
message = "SelectionPermutation x() :" & vbNewLine & _
        ShowPoly(x) & vbNewLine & vbNewLine
For i = 1 To n
    message = message & i & vbTab & Permutation(i + pLM1) & vbTab & x(Permutation(i + pLM1)) & vbNewLine
Next i
MsgBox message
'InsertionPermutation
Call InsertionPermutation(x(), Permutation())
message = "InsertionPermutation x() :" & vbNewLine & _
        ShowPoly(x) & vbNewLine & vbNewLine
For i = 1 To n
    message = message & i & vbTab & Permutation(i + pLM1) & vbTab & x(Permutation(i + pLM1)) & vbNewLine
Next i
MsgBox message
'ShellPermutation
Call ShellPermutation(x(), Permutation())
message = "ShellPermutation x() :" & vbNewLine & _
        ShowPoly(x) & vbNewLine & vbNewLine
For i = 1 To n
    message = message & i & vbTab & Permutation(i + pLM1) & vbTab & x(Permutation(i + pLM1)) & vbNewLine
Next i
MsgBox message
'HeapPermutation
Call HeapPermutation(x(), Permutation())
message = "HeapPermutation x() :" & vbNewLine & _
        ShowPoly(x) & vbNewLine & vbNewLine
For i = 1 To n
    message = message & i & vbTab & Permutation(i + pLM1) & vbTab & x(Permutation(i + pLM1)) & vbNewLine
Next i
MsgBox message
'MergePermutation
Call MergePermutation(x(), Permutation())
message = "MergePermutation x() :" & vbNewLine & _
        ShowPoly(x) & vbNewLine & vbNewLine
For i = 1 To n
    message = message & i & vbTab & Permutation(i + pLM1) & vbTab & x(Permutation(i + pLM1)) & vbNewLine
Next i
MsgBox message
'QuickPermutation
Call QuickPermutation(x(), Permutation())
message = "QuickPermutation x() :" & vbNewLine & _
        ShowPoly(x) & vbNewLine & vbNewLine
For i = 1 To n
    message = message & i & vbTab & Permutation(i + pLM1) & vbTab & x(Permutation(i + pLM1)) & vbNewLine
Next i
MsgBox message

End Sub
'x(LowerBound) stores minimum value
'x(UpperBound) stores maximum value
Private Sub BubbleSort(x() As Double)
Dim i As Long, j As Long
Dim ii As Long, jj As Long
Dim xLM1 As Long
Dim NumElement As Long
Dim tmp As Double

xLM1 = LBound(x) - 1&
NumElement = UBound(x) - xLM1

For i = 1& To NumElement - 1&
    ii = i + xLM1
    For j = i + 1& To NumElement
        jj = j + xLM1
        tmp = x(ii)
        If tmp > x(jj) Then
            x(ii) = x(jj)
            x(jj) = tmp
        End If
    Next j
Next i
End Sub
'Permutation(LowerBound) is the index for minimum value
'Permutation(UpperBound) is the index for maximum value
Private Sub BubblePermutation(x() As Double, Permutation() As Long)
Dim i As Long, j As Long
Dim ii As Long, jj As Long
Dim xLM1 As Long, pLM1 As Long
Dim NumElement As Long
Dim ltmp As Long
Dim tmp As Double
Dim xtmp() As Double

xLM1 = LBound(x) - 1&
pLM1 = LBound(Permutation) - 1&
NumElement = UBound(x) - xLM1
If UBound(Permutation) - pLM1 <> NumElement Then
    MsgBox "Size of parameter array ""Permutation"" is not" & _
           "equal to that of parameter array ""x""!"
    Exit Sub
End If

ReDim xtmp(1& To NumElement)
For i = 1& To NumElement
    ii = i + xLM1
    Permutation(i + pLM1) = ii
    xtmp(i) = x(ii)
Next i

For i = 1& To NumElement - 1&
    ii = i + pLM1
    For j = i + 1& To NumElement
        If xtmp(i) > xtmp(j) Then
            tmp = xtmp(i)
            xtmp(i) = xtmp(j)
            xtmp(j) = tmp
            jj = j + pLM1
            ltmp = Permutation(ii)
            Permutation(ii) = Permutation(jj)
            Permutation(jj) = ltmp
        End If
    Next j
Next i
End Sub
'x(LowerBound) stores minimum value
'x(UpperBound) stores maximum value
Private Sub SelectionSort(x() As Double)
Dim i As Long, j As Long
Dim ii As Long, jj As Long
Dim MinIndex As Long
Dim xLM1 As Long
Dim NumElement As Long
Dim tmp As Double

xLM1 = LBound(x) - 1&
NumElement = UBound(x) - xLM1

For i = 1& To NumElement - 1&
    ii = i + xLM1
    MinIndex = ii
    tmp = x(MinIndex)
    For j = i + 1& To NumElement
        jj = j + xLM1
        If x(jj) < tmp Then
            MinIndex = jj
            tmp = x(MinIndex)
        End If
    Next j
    If MinIndex > ii Then
        x(MinIndex) = x(ii)
        x(ii) = tmp
    End If
Next i
End Sub
'Permutation(LowerBound) is the index for minimum value
'Permutation(UpperBound) is the index for maximum value
Private Sub SelectionPermutation(x() As Double, Permutation() As Long)
Dim i As Long, j As Long
Dim ii As Long, jj As Long
Dim MinIndex As Long
Dim xLM1 As Long, pLM1 As Long
Dim NumElement As Long
Dim ltmp As Long
Dim tmp As Double
Dim xtmp() As Double

xLM1 = LBound(x) - 1&
pLM1 = LBound(Permutation) - 1&
NumElement = UBound(x) - xLM1
If UBound(Permutation) - pLM1 <> NumElement Then
    MsgBox "Size of parameter array ""Permutation"" is not" & _
           "equal to that of parameter array ""x""!"
    Exit Sub
End If

ReDim xtmp(1& To NumElement)
For i = 1& To NumElement
    ii = i + xLM1
    Permutation(i + pLM1) = ii
    xtmp(i) = x(ii)
Next i

For i = 1& To NumElement - 1&
    MinIndex = i
    tmp = xtmp(i)
    For j = i + 1& To NumElement
        If xtmp(j) < tmp Then
            MinIndex = j
            tmp = xtmp(MinIndex)
        End If
    Next j
    If MinIndex > i Then
        tmp = xtmp(i)
        xtmp(i) = xtmp(MinIndex)
        xtmp(MinIndex) = tmp
        ii = i + pLM1
        jj = MinIndex + pLM1
        ltmp = Permutation(ii)
        Permutation(ii) = Permutation(jj)
        Permutation(jj) = ltmp
    End If
Next i
End Sub
'x(LowerBound) stores minimum value
'x(UpperBound) stores maximum value
Private Sub InsertionSort(x() As Double)
Dim i As Long, j As Long
Dim ii As Long, jj As Long, jjM1 As Long
Dim xLM1 As Long
Dim NumElement As Long
Dim tmp As Double

xLM1 = LBound(x) - 1&
NumElement = UBound(x) - xLM1

j = 1& + xLM1
For i = 2& To NumElement
    ii = i + xLM1
    tmp = x(ii)
    jj = ii
    Do While jj > j
        jjM1 = jj - 1&
        If x(jjM1) > tmp Then
            x(jj) = x(jjM1)
            jj = jjM1
        Else
            Exit Do
        End If
    Loop
    If jj < ii Then
        x(jj) = tmp
    End If
Next i
End Sub
'Permutation(LowerBound) is the index for minimum value
'Permutation(UpperBound) is the index for maximum value
Private Sub InsertionPermutation(x() As Double, Permutation() As Long)
Dim i As Long, j As Long
Dim ii As Long, jj As Long
Dim jM1 As Long
Dim xLM1 As Long, pLM1 As Long
Dim NumElement As Long
Dim ltmp As Long
Dim tmp As Double
Dim xtmp() As Double

xLM1 = LBound(x) - 1&
pLM1 = LBound(Permutation) - 1&
NumElement = UBound(x) - xLM1
If UBound(Permutation) - pLM1 <> NumElement Then
    MsgBox "Size of parameter array ""Permutation"" is not" & _
           "equal to that of parameter array ""x""!"
    Exit Sub
End If

ReDim xtmp(1& To NumElement)
For i = 1& To NumElement
    ii = i + xLM1
    Permutation(i + pLM1) = ii
    xtmp(i) = x(ii)
Next i

For i = 2& To NumElement
    tmp = xtmp(i)
    ltmp = Permutation(i + pLM1)
    j = i
    Do While j > 1&
        jM1 = j - 1&
        If xtmp(jM1) > tmp Then
            xtmp(j) = xtmp(jM1)
            Permutation(j + pLM1) = Permutation(jM1 + pLM1)
            j = jM1
        Else
            Exit Do
        End If
    Loop
    If j < i Then
        xtmp(j) = tmp
        Permutation(j + pLM1) = ltmp
    End If
Next i
End Sub
'x(LowerBound) stores minimum value
'x(UpperBound) stores maximum value
Private Sub ShellSort(x() As Double)
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long, jjMstepsize As Long
Dim stepsize As Long
Dim xLM1 As Long
Dim NumElement As Long
Dim FirstStepRatio As Double, StepRatio As Double
Dim tmp As Double

xLM1 = LBound(x) - 1&
NumElement = UBound(x) - xLM1

If NumElement <= 500& Then
    FirstStepRatio = 0.1
    StepRatio = 0.4
ElseIf 500 < NumElement And NumElement < 5000 Then
    FirstStepRatio = 0.2
    StepRatio = 0.3
ElseIf 5000 <= NumElement Then
    FirstStepRatio = 0.2
    StepRatio = 0.4
End If

stepsize = Int(NumElement * FirstStepRatio)
If stepsize < 1& Then stepsize = 1&
Do
    For i = 1& To stepsize
        ii = i + xLM1
        For j = i + stepsize To NumElement Step stepsize
            jj = j + xLM1
            k = jj
            tmp = x(k)
            Do While jj > ii
                jjMstepsize = jj - stepsize
                If x(jjMstepsize) > tmp Then
                    x(jj) = x(jjMstepsize)
                    jj = jjMstepsize
                Else
                    Exit Do
                End If
            Loop
            If jj < k Then
                x(jj) = tmp
            End If
        Next j
    Next i
    If stepsize = 1& Then Exit Do
    stepsize = Int(stepsize * StepRatio)
    If stepsize < 1& Then stepsize = 1&
Loop
End Sub
'Permutation(LowerBound) is the index for minimum value
'Permutation(UpperBound) is the index for maximum value
Private Sub ShellPermutation(x() As Double, Permutation() As Long)
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long
Dim jMstepsize As Long
Dim stepsize As Long
Dim xLM1 As Long, pLM1 As Long
Dim NumElement As Long
Dim FirstStepRatio As Double, StepRatio As Double
Dim ltmp As Long
Dim tmp As Double
Dim xtmp() As Double

xLM1 = LBound(x) - 1&
pLM1 = LBound(Permutation) - 1&
NumElement = UBound(x) - xLM1
If UBound(Permutation) - pLM1 <> NumElement Then
    MsgBox "Size of parameter array ""Permutation"" is not" & _
           "equal to that of parameter array ""x""!"
    Exit Sub
End If

ReDim xtmp(1& To NumElement)
For i = 1& To NumElement
    ii = i + xLM1
    Permutation(i + pLM1) = ii
    xtmp(i) = x(ii)
Next i

If NumElement <= 500& Then
    FirstStepRatio = 0.1
    StepRatio = 0.4
ElseIf 500 < NumElement And NumElement < 5000 Then
    FirstStepRatio = 0.2
    StepRatio = 0.3
ElseIf 5000 <= NumElement Then
    FirstStepRatio = 0.2
    StepRatio = 0.4
End If

stepsize = Int(NumElement * FirstStepRatio)
If stepsize < 1& Then stepsize = 1&
Do
    For i = 1& To stepsize
        For j = i + stepsize To NumElement Step stepsize
            k = j
            tmp = xtmp(k)
            ltmp = Permutation(k + pLM1)
            Do While j > i
                jMstepsize = j - stepsize
                If xtmp(jMstepsize) > tmp Then
                    xtmp(j) = xtmp(jMstepsize)
                    Permutation(j + pLM1) = Permutation(jMstepsize + pLM1)
                    j = jMstepsize
                Else
                    Exit Do
                End If
            Loop
            If j < k Then
                xtmp(j) = tmp
                Permutation(j + pLM1) = ltmp
            End If
        Next j
    Next i
    If stepsize = 1& Then Exit Do
    stepsize = Int(stepsize * StepRatio)
    If stepsize < 1& Then stepsize = 1&
Loop
End Sub
'x(LowerBound) stores minimum value
'x(UpperBound) stores maximum value
Private Sub HeapSort(x() As Double)
Dim lower As Long, j As Long
Dim upper As Long
Dim parent As Long, child As Long, childP1 As Long
Dim xLM1 As Long
Dim tmp As Double, tmp1 As Double

xLM1 = LBound(x) - 1&
upper = UBound(x) - xLM1
lower = Int(upper / 2&)
Do
    If (lower > 1&) Then
'construct heap
        lower = lower - 1&
        tmp = x(lower + xLM1)
    Else
'reconstruct heap
'now lower = 1
        If (upper = 1&) Then Exit Sub
        j = upper + xLM1
        tmp = x(j)
        x(j) = x(1& + xLM1)
        upper = upper - 1&
    End If
    
    parent = lower
    child = lower * 2&
    childP1 = child + 1&
    Do While (child <= upper)
        j = child
        If (childP1 <= upper) Then
            If (x(childP1 + xLM1) > x(child + xLM1)) Then
                j = childP1
            End If
        End If
        tmp1 = x(j + xLM1)
        If (tmp1 > tmp) Then
            x(parent + xLM1) = tmp1
            parent = j
            child = parent * 2&
            childP1 = child + 1&
        Else
            Exit Do
        End If
    Loop
    x(parent + xLM1) = tmp
Loop
End Sub
'Permutation(LowerBound) is the index for minimum value
'Permutation(UpperBound) is the index for maximum value
Private Sub HeapPermutation(x() As Double, Permutation() As Long)
Dim lower As Long, j As Long
Dim upper As Long
Dim parent As Long, child As Long, childP1 As Long
Dim xLM1 As Long, pLM1 As Long
Dim ltmp As Long, ltmp1 As Long
Dim tmp As Double, tmp1 As Double
Dim xtmp() As Double

xLM1 = LBound(x) - 1&
pLM1 = LBound(Permutation) - 1&
upper = UBound(x) - xLM1
lower = Int(upper / 2&)
If UBound(Permutation) - pLM1 <> upper Then
    MsgBox "Size of parameter array ""Permutation"" is not" & _
           "equal to that of parameter array ""x""!"
    Exit Sub
End If

ReDim xtmp(1& To upper)
For j = 1& To upper
    ltmp = j + xLM1
    Permutation(j + pLM1) = ltmp
    xtmp(j) = x(ltmp)
Next j

Do
    If (lower > 1&) Then
'construct heap
        lower = lower - 1&
        tmp = xtmp(lower)
        ltmp = Permutation(lower + pLM1)
    Else
'reconstruct heap
'now lower = 1
        If (upper = 1&) Then Exit Sub
        j = upper
        tmp = xtmp(j)
        ltmp = Permutation(j + pLM1)
        xtmp(j) = xtmp(1&)
        Permutation(j + pLM1) = Permutation(1& + pLM1)
        upper = upper - 1&
    End If
    
    parent = lower
    child = lower * 2&
    childP1 = child + 1&
    Do While (child <= upper)
        j = child
        If (childP1 <= upper) Then
            If (xtmp(childP1) > xtmp(child)) Then
                j = childP1
            End If
        End If
        tmp1 = xtmp(j)
        ltmp1 = Permutation(j + pLM1)
        If (tmp1 > tmp) Then
            xtmp(parent) = tmp1
            Permutation(parent + pLM1) = ltmp1
            parent = j
            child = parent * 2&
            childP1 = child + 1&
        Else
            Exit Do
        End If
    Loop
    xtmp(parent) = tmp
    Permutation(parent + pLM1) = ltmp
Loop
End Sub
'x(LowerBound) stores minimum value
'x(UpperBound) stores maximum value
Public Sub MergeSort(x() As Double)

Call MergeSort_(x(), LBound(x), UBound(x))
End Sub
'
Private Sub MergeSort_(x() As Double, lower As Long, upper As Long)

If (upper > lower) Then
    Dim middle As Long
    middle = Int(CDbl(upper + lower) * 0.5)
    Call MergeSort_(x(), lower, middle)
    Call MergeSort_(x(), middle + 1&, upper)
    Call Merge_(x(), lower, middle, upper)
End If
End Sub
'
Private Sub Merge_(x() As Double, _
                   lower As Long, middle As Long, upper As Long)
Dim LowerStart As Long, LowerEnd As Long
Dim UpperStart As Long, UpperEnd As Long
Dim i As Long
Dim tmp As Double

LowerStart = lower
LowerEnd = middle
UpperStart = middle + 1&
UpperEnd = upper

Do While ((LowerStart <= LowerEnd) And _
          (UpperStart <= UpperEnd))
    tmp = x(UpperStart)
    If (tmp < x(LowerStart)) Then
        For i = UpperStart To LowerStart + 1& Step -1&
            x(i) = x(i - 1&)
        Next i
        x(LowerStart) = tmp
        LowerStart = LowerStart + 1&
        LowerEnd = LowerEnd + 1&
        UpperStart = UpperStart + 1&
    Else
        LowerStart = LowerStart + 1&
    End If
Loop
End Sub
'Permutation(LowerBound) is the index for minimum value
'Permutation(UpperBound) is the index for maximum value
Public Sub MergePermutation(x() As Double, Permutation() As Long)
Dim i As Long, ii As Long
Dim xLM1 As Long, pLM1 As Long
Dim NumElement As Long
Dim xtmp() As Double

xLM1 = LBound(x) - 1&
pLM1 = LBound(Permutation) - 1&
NumElement = UBound(x) - xLM1
If UBound(Permutation) - pLM1 <> NumElement Then
    MsgBox "Size of parameter array ""Permutation"" is not" & _
           "equal to that of parameter array ""x""!"
    Exit Sub
End If

ReDim xtmp(1& + xLM1 To NumElement + xLM1)
For i = 1& To NumElement
    ii = i + xLM1
    Permutation(i + pLM1) = ii
    xtmp(ii) = x(ii)
Next i

Call MergePermutation_(xtmp(), 1& + xLM1, NumElement + xLM1, Permutation(), pLM1 - xLM1)
End Sub
'
Private Sub MergePermutation_(xtmp() As Double, lower As Long, upper As Long, _
                              Permutation() As Long, xTmpToPerm As Long)

If (upper > lower) Then
    Dim middle As Long
    middle = Int(CDbl(upper + lower) * 0.5)
    Call MergePermutation_(xtmp(), lower, middle, Permutation(), xTmpToPerm)
    Call MergePermutation_(xtmp(), middle + 1&, upper, Permutation(), xTmpToPerm)
    Call MergeP_(xtmp(), lower, middle, upper, Permutation(), xTmpToPerm)
End If
End Sub
'
Private Sub MergeP_(xtmp() As Double, _
                    lower As Long, middle As Long, upper As Long, _
                    Permutation() As Long, xTmpToPerm As Long)
Dim LowerStart As Long, LowerEnd As Long
Dim UpperStart As Long, UpperEnd As Long
Dim i As Long, j As Long
Dim ltmp As Long
Dim tmp As Double

LowerStart = lower
LowerEnd = middle
UpperStart = middle + 1&
UpperEnd = upper

Do While ((LowerStart <= LowerEnd) And _
          (UpperStart <= UpperEnd))
    tmp = xtmp(UpperStart)
    ltmp = Permutation(UpperStart + xTmpToPerm)
    If (tmp < xtmp(LowerStart)) Then
        For i = UpperStart To LowerStart + 1& Step -1&
            j = i - 1&
            xtmp(i) = xtmp(j)
            Permutation(i + xTmpToPerm) = Permutation(j + xTmpToPerm)
        Next i
        xtmp(LowerStart) = tmp
        Permutation(LowerStart + xTmpToPerm) = ltmp
        LowerStart = LowerStart + 1&
        LowerEnd = LowerEnd + 1&
        UpperStart = UpperStart + 1&
    Else
        LowerStart = LowerStart + 1&
    End If
Loop
End Sub
'x(LowerBound) stores minimum value
'x(UpperBound) stores maximum value
Public Sub QuickSort(x() As Double)

Randomize timer
Call QuickSort_(x(), LBound(x), UBound(x))
End Sub
'
Private Sub QuickSort_(x() As Double, lower As Long, upper As Long)
Dim LowerIndex As Long, UpperIndex As Long
Dim tmpIndex As Long, PivotIndex As Long
Dim pivot As Double, tmp As Double

LowerIndex = lower
UpperIndex = upper
PivotIndex = Int((upper - lower + 1&) * Rnd + lower)
pivot = x(PivotIndex)
Do While (LowerIndex < UpperIndex)
    Do While ((x(UpperIndex) >= pivot) And _
              (LowerIndex < UpperIndex))
        UpperIndex = UpperIndex - 1&
    Loop
    Do While ((x(LowerIndex) <= pivot) And _
              (LowerIndex < UpperIndex))
        LowerIndex = LowerIndex + 1&
    Loop
    If (LowerIndex < UpperIndex) Then
        tmp = x(UpperIndex)
        x(UpperIndex) = x(LowerIndex)
        x(LowerIndex) = tmp
        UpperIndex = UpperIndex - 1&
        LowerIndex = LowerIndex + 1&
    End If
Loop

If LowerIndex = UpperIndex Then
    tmp = x(LowerIndex)
    If (tmp < pivot) Then
        If PivotIndex > LowerIndex Then
            tmpIndex = LowerIndex + 1&
            x(PivotIndex) = x(tmpIndex)
            x(tmpIndex) = pivot
            If (lower < LowerIndex) Then
                Call QuickSort_(x, lower, LowerIndex)
            End If
            tmpIndex = LowerIndex + 2&
            If (upper > tmpIndex) Then
                Call QuickSort_(x, tmpIndex, upper)
            End If
            Exit Sub
        ElseIf PivotIndex < LowerIndex Then
            x(PivotIndex) = tmp
            x(LowerIndex) = pivot
        Else

        End If
    ElseIf (tmp > pivot) Then
        If PivotIndex > LowerIndex Then
            x(PivotIndex) = tmp
            x(LowerIndex) = pivot
        ElseIf PivotIndex < LowerIndex Then
            tmpIndex = LowerIndex - 1&
            x(PivotIndex) = x(tmpIndex)
            x(tmpIndex) = pivot
            tmpIndex = LowerIndex - 2&
            If (lower < tmpIndex) Then
                Call QuickSort_(x, lower, tmpIndex)
            End If
            If (upper > LowerIndex) Then
                Call QuickSort_(x, LowerIndex, upper)
            End If
            Exit Sub
        Else

        End If
    Else

    End If
    tmpIndex = LowerIndex - 1&
    If (lower < tmpIndex) Then
        Call QuickSort_(x, lower, tmpIndex)
    End If
    tmpIndex = LowerIndex + 1&
    If (upper > tmpIndex) Then
        Call QuickSort_(x, tmpIndex, upper)
    End If
Else
    If (lower < UpperIndex) Then Call QuickSort_(x, lower, UpperIndex)
    If (upper > LowerIndex) Then Call QuickSort_(x, LowerIndex, upper)
End If
End Sub
'Permutation(LowerBound) is the index for minimum value
'Permutation(UpperBound) is the index for maximum value
Public Sub QuickPermutation(x() As Double, Permutation() As Long)
Dim i As Long, ii As Long
Dim xLM1 As Long, pLM1 As Long
Dim NumElement As Long
Dim xtmp() As Double

xLM1 = LBound(x) - 1&
pLM1 = LBound(Permutation) - 1&
NumElement = UBound(x) - xLM1
If UBound(Permutation) - pLM1 <> NumElement Then
    MsgBox "Size of parameter array ""Permutation"" is not" & _
           "equal to that of parameter array ""x""!"
    Exit Sub
End If

ReDim xtmp(1& + xLM1 To NumElement + xLM1)
For i = 1& To NumElement
    ii = i + xLM1
    Permutation(i + pLM1) = ii
    xtmp(ii) = x(ii)
Next i

Randomize timer
Call QuickSPermutation_(xtmp(), 1& + xLM1, NumElement + xLM1, Permutation(), pLM1 - xLM1)
End Sub
'
Private Sub QuickSPermutation_(xtmp() As Double, lower As Long, upper As Long, _
                               Permutation() As Long, xTmpToPerm As Long)
Dim LowerIndex As Long, UpperIndex As Long
Dim tmpIndex As Long, tmp1Index As Long
Dim PivotIndex As Long
Dim pivot As Double, tmp As Double

LowerIndex = lower
UpperIndex = upper
PivotIndex = Int((upper - lower + 1&) * Rnd + lower)
pivot = xtmp(PivotIndex)
Do While (LowerIndex < UpperIndex)
    Do While ((xtmp(UpperIndex) >= pivot) And _
              (LowerIndex < UpperIndex))
        UpperIndex = UpperIndex - 1&
    Loop
    Do While ((xtmp(LowerIndex) <= pivot) And _
              (LowerIndex < UpperIndex))
        LowerIndex = LowerIndex + 1&
    Loop
    If (LowerIndex < UpperIndex) Then
        tmp = xtmp(UpperIndex)
        xtmp(UpperIndex) = xtmp(LowerIndex)
        xtmp(LowerIndex) = tmp
        tmpIndex = Permutation(UpperIndex + xTmpToPerm)
        Permutation(UpperIndex + xTmpToPerm) = Permutation(LowerIndex + xTmpToPerm)
        Permutation(LowerIndex + xTmpToPerm) = tmpIndex
        UpperIndex = UpperIndex - 1&
        LowerIndex = LowerIndex + 1&
    End If
Loop

If LowerIndex = UpperIndex Then
    tmp = xtmp(LowerIndex)
    If (tmp < pivot) Then
        If PivotIndex > LowerIndex Then
            tmpIndex = LowerIndex + 1&
            xtmp(PivotIndex) = xtmp(tmpIndex)
            xtmp(tmpIndex) = pivot
            tmp1Index = Permutation(PivotIndex + xTmpToPerm)
            Permutation(PivotIndex + xTmpToPerm) = Permutation(tmpIndex + xTmpToPerm)
            Permutation(tmpIndex + xTmpToPerm) = tmp1Index
            If (lower < LowerIndex) Then
                Call QuickSPermutation_(xtmp(), lower, LowerIndex, Permutation(), xTmpToPerm)
            End If
            tmpIndex = LowerIndex + 2&
            If (upper > tmpIndex) Then
                Call QuickSPermutation_(xtmp(), tmpIndex, upper, Permutation(), xTmpToPerm)
            End If
            Exit Sub
        ElseIf PivotIndex < LowerIndex Then
            xtmp(PivotIndex) = tmp
            xtmp(LowerIndex) = pivot
            tmp1Index = Permutation(PivotIndex + xTmpToPerm)
            Permutation(PivotIndex + xTmpToPerm) = Permutation(LowerIndex + xTmpToPerm)
            Permutation(LowerIndex + xTmpToPerm) = tmp1Index
        Else

        End If
    ElseIf (tmp > pivot) Then
        If PivotIndex > LowerIndex Then
            xtmp(PivotIndex) = tmp
            xtmp(LowerIndex) = pivot
            tmp1Index = Permutation(PivotIndex + xTmpToPerm)
            Permutation(PivotIndex + xTmpToPerm) = Permutation(LowerIndex + xTmpToPerm)
            Permutation(LowerIndex + xTmpToPerm) = tmp1Index
        ElseIf PivotIndex < LowerIndex Then
            tmpIndex = LowerIndex - 1&
            xtmp(PivotIndex) = xtmp(tmpIndex)
            xtmp(tmpIndex) = pivot
            tmp1Index = Permutation(PivotIndex + xTmpToPerm)
            Permutation(PivotIndex + xTmpToPerm) = Permutation(tmpIndex + xTmpToPerm)
            Permutation(tmpIndex + xTmpToPerm) = tmp1Index
            tmpIndex = LowerIndex - 2&
            If (lower < tmpIndex) Then
                Call QuickSPermutation_(xtmp(), lower, tmpIndex, Permutation(), xTmpToPerm)
            End If
            If (upper > LowerIndex) Then
                Call QuickSPermutation_(xtmp(), LowerIndex, upper, Permutation(), xTmpToPerm)
            End If
            Exit Sub
        Else

        End If
    Else

    End If
    tmpIndex = LowerIndex - 1&
    If (lower < tmpIndex) Then
        Call QuickSPermutation_(xtmp(), lower, tmpIndex, Permutation(), xTmpToPerm)
    End If
    tmpIndex = LowerIndex + 1&
    If (upper > tmpIndex) Then
        Call QuickSPermutation_(xtmp(), tmpIndex, upper, Permutation(), xTmpToPerm)
    End If
Else
    If (lower < UpperIndex) Then Call QuickSPermutation_(xtmp, lower, UpperIndex, Permutation(), xTmpToPerm)
    If (upper > LowerIndex) Then Call QuickSPermutation_(xtmp, LowerIndex, upper, Permutation(), xTmpToPerm)
End If
End Sub

