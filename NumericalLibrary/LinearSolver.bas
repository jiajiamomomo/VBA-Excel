Attribute VB_Name = "LinearSolver"
Option Explicit
'
'
Public Sub TestLinearSolver()
Dim i As Integer, j As Integer, k As Integer
Const A1LM1 As Integer = 5
Const A2LM1 As Integer = 4
Const bLM1 As Integer = 3
Const xLM1 As Integer = 2
Const n As Integer = 50
Dim tmp As Double
Dim TimerStart As Double, TimerEnd As Double
Dim A(1 + A1LM1 To n + A1LM1, 1 + A2LM1 To n + A2LM1) As Double
Dim B(1 + bLM1 To n + bLM1) As Double
Dim x(1 + xLM1 To n + xLM1) As Double
Dim sol(1 To n) As Double
Dim inv As Variant
Dim xx As Variant
Dim sh As Worksheet

ReDim inv(1 To n, 1 To n)
ReDim xx(1 To n)
Randomize timer

For i = 1 To n
    A(i + A1LM1, i + A2LM1) = Rnd
    For j = i + 1 To n
        tmp = Sgn(Rnd - 0.5) * Rnd ^ (12# * (Rnd - 0.5))
        A(i + A1LM1, j + A2LM1) = tmp
        A(j + A1LM1, i + A2LM1) = tmp
    Next j
    sol(i) = Sgn(Rnd - 0.5) * Rnd '^ (10# * (Rnd - 0.5))
Next i
For i = 1 To n
    tmp = 0#
    For j = 1 To n
        tmp = tmp + A(i + A1LM1, j + A2LM1) * sol(j)
    Next j
    B(i + bLM1) = tmp
Next i

'LULinearSolver
'for general matrix
'solved solution is affected by the vaue of the coefficient
'of the matrix
Application.ScreenUpdating = False
TimerStart = timer
For k = 1 To 100
    Call LULinearSolver(A(), B(), x())
Next k
TimerEnd = timer
MsgBox "LULinearSolver: " & (TimerEnd - TimerStart) & "secs"

TimerStart = timer
For k = 1 To 100
    inv = Application.WorksheetFunction.MInverse(A())
    For i = 1 To n
        tmp = 0#
        For j = 1 To n
            tmp = tmp + inv(i, j) * B(j + bLM1)
        Next j
        xx(i) = tmp
    Next i
Next k
TimerEnd = timer
MsgBox "MInverse: " & (TimerEnd - TimerStart) & "secs"

Set sh = ThisWorkbook.Worksheets("LULinearSolver")
sh.Cells.Clear

sh.Cells(1, 1) = "LULinearSolver"
sh.Range(sh.Cells(2, 1), sh.Cells(1 + n, 1)).Value = Application.WorksheetFunction.Transpose(x)

sh.Cells(1, 2) = "sol"
sh.Range(sh.Cells(2, 2), sh.Cells(1 + n, 2)).Value = Application.WorksheetFunction.Transpose(sol)

For i = 1 To n
    tmp = B(i + bLM1)
    For j = 1 To n
        tmp = tmp - A(i + A1LM1, j + A2LM1) * x(j + xLM1)
    Next j
    xx(i) = tmp
Next i

sh.Cells(1, 3) = "Residue %"
sh.Range(sh.Cells(2, 3), sh.Cells(1 + n, 3)).Value = Application.WorksheetFunction.Transpose(xx)

'LUSymmetricLinearSolver
'for symmetric matrix only
TimerStart = timer
For k = 1 To 100
    Call LUSymmetricLinearSolver(A(), B(), x())
Next k
TimerEnd = timer
MsgBox "LUSymmetricLinearSolver: " & (TimerEnd - TimerStart) & "secs"

TimerStart = timer
For k = 1 To 100
    inv = Application.WorksheetFunction.MInverse(A())
    For i = 1 To n
        tmp = 0#
        For j = 1 To n
            tmp = tmp + inv(i, j) * B(j + bLM1)
        Next j
        xx(i) = tmp
    Next i
Next k
TimerEnd = timer
MsgBox "MInverse: " & (TimerEnd - TimerStart) & "secs"

Set sh = ThisWorkbook.Worksheets("LUSymmetric")
sh.Cells.Clear

sh.Cells(1, 1) = "LUSymmetricLinearSolver"
sh.Range(sh.Cells(2, 1), sh.Cells(1 + n, 1)).Value = Application.WorksheetFunction.Transpose(x)

sh.Cells(1, 2) = "sol"
sh.Range(sh.Cells(2, 2), sh.Cells(1 + n, 2)).Value = Application.WorksheetFunction.Transpose(sol)

For i = 1 To n
    tmp = B(i + bLM1)
    For j = 1 To n
        tmp = tmp - A(i + A1LM1, j + A2LM1) * x(j + xLM1)
    Next j
    xx(i) = tmp
Next i

sh.Cells(1, 3) = "Residue %"
sh.Range(sh.Cells(2, 3), sh.Cells(1 + n, 3)).Value = Application.WorksheetFunction.Transpose(xx)

'HouseholderLinearSolver
'consume more time
'solved solution is almost not affected by the vaue of
'the coefficient of the matrix
TimerStart = timer
For k = 1 To 100
    Call HouseholderLinearSolver(A(), B(), x())
Next k
TimerEnd = timer
MsgBox "HouseholderLinearSolver: " & (TimerEnd - TimerStart) & "secs"

TimerStart = timer
For k = 1 To 100
    inv = Application.WorksheetFunction.MInverse(A())
    For i = 1 To n
        tmp = 0#
        For j = 1 To n
            tmp = tmp + inv(i, j) * B(j + bLM1)
        Next j
        xx(i) = tmp
    Next i
Next k
TimerEnd = timer
MsgBox "MInverse: " & (TimerEnd - TimerStart) & "secs"

Set sh = ThisWorkbook.Worksheets("Householder")
sh.Cells.Clear

sh.Cells(1, 1) = "HouseholderLinearSolver"
sh.Range(sh.Cells(2, 1), sh.Cells(1 + n, 1)).Value = Application.WorksheetFunction.Transpose(x)

sh.Cells(1, 2) = "sol"
sh.Range(sh.Cells(2, 2), sh.Cells(1 + n, 2)).Value = Application.WorksheetFunction.Transpose(sol)

For i = 1 To n
    tmp = B(i + bLM1)
    For j = 1 To n
        tmp = tmp - A(i + A1LM1, j + A2LM1) * x(j + xLM1)
    Next j
    xx(i) = tmp
Next i

sh.Cells(1, 3) = "Residue %"
sh.Range(sh.Cells(2, 3), sh.Cells(1 + n, 3)).Value = Application.WorksheetFunction.Transpose(xx)

'LUPivotLinearSolver
'
TimerStart = timer
For k = 1 To 100
    Call LUPivotLinearSolver(A(), B(), x())
Next k
TimerEnd = timer
MsgBox "LUPivotLinearSolver: " & (TimerEnd - TimerStart) & "secs"

TimerStart = timer
For k = 1 To 100
    inv = Application.WorksheetFunction.MInverse(A())
    For i = 1 To n
        tmp = 0#
        For j = 1 To n
            tmp = tmp + inv(i, j) * B(j + bLM1)
        Next j
        xx(i) = tmp
    Next i
Next k
TimerEnd = timer
MsgBox "MInverse: " & (TimerEnd - TimerStart) & "secs"

Set sh = ThisWorkbook.Worksheets("LUPivot")
sh.Cells.Clear

sh.Cells(1, 1) = "LUPivotLinearSolver"
sh.Range(sh.Cells(2, 1), sh.Cells(1 + n, 1)).Value = Application.WorksheetFunction.Transpose(x)

sh.Cells(1, 2) = "sol"
sh.Range(sh.Cells(2, 2), sh.Cells(1 + n, 2)).Value = Application.WorksheetFunction.Transpose(sol)

For i = 1 To n
    tmp = B(i + bLM1)
    For j = 1 To n
        tmp = tmp - A(i + A1LM1, j + A2LM1) * x(j + xLM1)
    Next j
    xx(i) = tmp
Next i

sh.Cells(1, 3) = "Residue %"
sh.Range(sh.Cells(2, 3), sh.Cells(1 + n, 3)).Value = Application.WorksheetFunction.Transpose(xx)



Application.ScreenUpdating = True

End Sub
'A * x = b
'A = L * U
'L * U * x = b
'L * (U * x) = b
'L * y = b
'U * x = y
'
'     j=1  j=2  j=3  .    j=n
'    +a11  0    0    .    0  + i=1
'    |a21  a22  0    .    0  | i=2
'L = |a31  a32  a33  .    0  | i=3
'    |.    .    .    .    .  | .
'    +an1  an2  an3  .    ann+ i=n
'
'     j=1  j=2  j=3  .    j=n
'    +1    a12  a13  .    a1n+ i=1
'    |0    1    a23  .    a2n| i=2
'U = |0    0    1    .    a3n| i=3
'    |.    .    .    .    .  | .
'    +0    0    0    .    1  + i=n
'
'combine matrix L and matrix U as matrix LU
'
'      j=1  j=2  j=3  .    j=n
'     +a11  a12  a13  .    a1n+ i=1
'     |a21  a22  a23  .    a2n| i=2
'LU = |a31  a32  a33  .    a3n| i=3
'     |.    .    .    .    .  | .
'     +an1  an2  an3  .    ann+ i=n
Public Sub LULinearSolver(A() As Double, B() As Double, x() As Double)
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim ii As Integer, jj As Integer
Dim A1LM1 As Integer, A2LM1 As Integer
Dim bLM1 As Integer, xLM1 As Integer
Dim UpperBound As Integer
Dim tmp As Double, tmp1 As Double
Dim LU() As Double
Dim y() As Double

A1LM1 = LBound(A(), 1) - 1
A2LM1 = LBound(A(), 2) - 1
bLM1 = LBound(B()) - 1
xLM1 = LBound(x()) - 1
UpperBound = UBound(A(), 1) - A1LM1

If UpperBound <> (UBound(A(), 2) - A2LM1) Then
    MsgBox "Parameter array A() is not square!"
End If
If UpperBound <> (UBound(B()) - bLM1) Then
    MsgBox "Number of element of parameter vector b() is not equal to that of A()!"
End If
If UpperBound <> (UBound(x()) - xLM1) Then
    MsgBox "Number of element of parameter vector x() is not equal to that of A()!"
End If

ReDim LU(1 To UpperBound, 1 To UpperBound)
ReDim y(1 To UpperBound)

'A = L * U
jj = 1 + A2LM1
For i = 1 To UpperBound
    LU(i, 1) = A(i + A1LM1, jj)
Next i
For k = 2 To UpperBound
'U
    i = k - 1
    ii = i + A1LM1
    tmp1 = LU(i, i)
    If tmp1 = 0# Then
        MsgBox "Singular Matrix" & vbNewLine & "Divided by zero", vbCritical
        Exit Sub
    End If
    tmp1 = 1# / tmp1
    For j = k To UpperBound
        tmp = A(ii, j + A2LM1)
        For l = 1 To i - 1
            tmp = tmp - LU(i, l) * LU(l, j)
        Next l
        LU(i, j) = tmp * tmp1
    Next j
'L
    j = k
    jj = j + A2LM1
    For i = k To UpperBound
        tmp = A(i + A1LM1, jj)
        For l = 1 To j - 1
            tmp = tmp - LU(i, l) * LU(l, j)
        Next l
        LU(i, j) = tmp
    Next i
Next k
'L * y = b
For i = 1 To UpperBound
    tmp = B(i + bLM1)
    For k = 1 To i - 1
        tmp = tmp - LU(i, k) * y(k)
    Next k
    y(i) = tmp / LU(i, i)
Next i
'U * x = y
For i = UpperBound To 1 Step -1
    tmp = y(i)
    For k = i + 1 To UpperBound
        tmp = tmp - LU(i, k) * x(k + xLM1)
    Next k
    x(i + xLM1) = tmp
Next i
End Sub
'A * x = b
'A = L * U
'L * U * x = b
'L * (U * x) = b
'L * y = b
'U * x = y
'
'     j=1  j=2  j=3  .    j=n
'    +a11  0    0    .    0  + i=1
'    |a21  a22  0    .    0  | i=2
'L = |a31  a32  a33  .    0  | i=3
'    |.    .    .    .    .  | .
'    +an1  an2  an3  .    ann+ i=n
'
'     j=1  j=2  j=3  .    j=n
'    +1    a12  a13  .    a1n+ i=1
'    |0    1    a23  .    a2n| i=2
'U = |0    0    1    .    a3n| i=3
'    |.    .    .    .    .  | .
'    +0    0    0    .    1  + i=n
'
'combine matrix L and matrix U as matrix LU
'
'      j=1  j=2  j=3  .    j=n
'     +a11  a12  a13  .    a1n+ i=1
'     |a21  a22  a23  .    a2n| i=2
'LU = |a31  a32  a33  .    a3n| i=3
'     |.    .    .    .    .  | .
'     +an1  an2  an3  .    ann+ i=n
Public Sub LUPivotLinearSolver(A() As Double, B() As Double, x() As Double)
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim ii As Integer, jj As Integer
Dim LA1 As Integer, LA2 As Integer, Lb As Integer, Lx As Integer
Dim UA1 As Integer, UA2 As Integer, Ub As Integer, Ux As Integer
Dim A1LM1 As Integer, A2LM1 As Integer
Dim bLM1 As Integer, xLM1 As Integer
Dim UpperBound As Integer
Dim tmp As Double, tmp1 As Double
Dim MaxElement As Double
Dim LU() As Double
Dim y() As Double

LA1 = LBound(A(), 1)
UA1 = UBound(A(), 1)
LA2 = LBound(A(), 2)
UA2 = UBound(A(), 2)
Lb = LBound(B())
Ub = UBound(B())
Lx = LBound(x())
Ux = UBound(x())

If (UA1 - LA1) <> (UA2 - LA2) Then
    MsgBox "Parameter array A() is not square!"
End If
If (UA1 - LA1) <> (Ub - Lb) Then
    MsgBox "Number of element of parameter vector b() is not equal to that of A()!"
End If
If (UA1 - LA1) <> (Ux - Lx) Then
    MsgBox "Number of element of parameter vector x() is not equal to that of A()!"
End If

A1LM1 = LA1 - 1
A2LM1 = LA2 - 1
bLM1 = Lb - 1
xLM1 = Lx - 1
UpperBound = Ub - bLM1

ReDim LU(1 To UpperBound, 1 To UpperBound)
ReDim y(1 To UpperBound)
ReDim Permutation(1 To UpperBound)

For i = 1 To UpperBound
    Permutation(i) = i
Next i

'Find maximum in the j-th column
For j = 1 To UpperBound - 1
    jj = j + A2LM1
'k is the index of max element
    MaxElement = Abs(A(j + A1LM1, jj))
    k = j
    For i = j + 1 To UpperBound
        tmp = Abs(A(i + A1LM1, jj))
        If tmp > MaxElement Then
            k = i
            MaxElement = tmp
        End If
    Next i
    If k > j Then
'swap
        ii = k + bLM1
        jj = j + bLM1
        tmp = B(ii)
        B(ii) = B(jj)
        B(jj) = tmp
'swap row
        ii = k + A1LM1
        jj = j + A1LM1
        For l = 1 To UpperBound
            i = l + A2LM1
            tmp = A(ii, i)
            A(ii, i) = A(jj, i)
            A(jj, i) = tmp
        Next l
    End If
Next j

'A = L * U
jj = 1 + A2LM1
For i = 1 To UpperBound
    LU(i, 1) = A(i + A1LM1, jj)
Next i
For k = 2 To UpperBound
'U
    i = k - 1
    ii = i + A1LM1
    tmp1 = LU(i, i)
    If tmp1 = 0# Then
        MsgBox "Singular Matrix" & vbNewLine & "Divided by zero", vbCritical
        Exit Sub
    End If
    tmp1 = 1# / tmp1
    For j = k To UpperBound
        tmp = A(ii, j + A2LM1)
        For l = 1 To i - 1
            tmp = tmp - LU(i, l) * LU(l, j)
        Next l
        tmp = tmp * tmp1
        LU(i, j) = tmp
    Next j
'L
    j = k
    jj = j + A2LM1
    For i = k To UpperBound
        tmp = A(i + A1LM1, jj)
        For l = 1 To j - 1
            tmp = tmp - LU(i, l) * LU(l, j)
        Next l
        LU(i, j) = tmp
    Next i
Next k
'L * y = b
For i = 1 To UpperBound
    tmp = B(i + bLM1)
    For k = 1 To i - 1
        tmp = tmp - LU(i, k) * y(k)
    Next k
    tmp = tmp / LU(i, i)
    y(i) = tmp
Next i
'U * x = y
For i = UpperBound To 1 Step -1
    tmp = y(i)
    For k = i + 1 To UpperBound
        tmp = tmp - LU(i, k) * x(k + xLM1)
    Next k
    x(i + xLM1) = tmp
Next i
End Sub
'A * x = b
'For symmetric matrix only!!!
'A = transpose(A)
'A = transpose(U) * D * U
'transpose(U) * D * U * x = b
'transpose(U) * (D * (U * x)) = b
'transpose(U) * z = b
'D * y = z
'U * x = y
'
'     j=1  j=2  j=3  .    j=n
'    +1    a12  a13  .    a1n+ i=1
'    |0    1    a23  .    a2n| i=2
'U = |0    0    1    .    a3n| i=3
'    |.    .    .    .    .  | .
'    +0    0    0    .    1  + i=n
'
'                j=1  j=2  j=3  .    j=n
'               +1    0    0    .    0  + i=1
'               |a12  1    0    .    0  | i=2
'transpose(U) = |a13  a23  1    .    0  | i=3
'               |.    .    .    .    .  | .
'               +a1n  a2n  a3n  .    1  + i=n
'
'     j=1  j=2  j=3  .    j=n
'    +a11  0    0    .    0  + i=1
'    |0    a22  0    .    0  | i=2
'D = |0    0    a33  .    0  | i=3
'    |.    .    .    .    .  | .
'    +0    0    0    .    ann+ i=n
'
'combine matrix transpose(U),matrix D and matrix U as
'matrix LU
'
'      j=1  j=2  j=3  .    j=n
'     +a11  a12  a13  .    a1n+ i=1
'     |a12  a22  a23  .    a2n| i=2
'LU = |a13  a23  a33  .    a3n| i=3
'     |.    .    .    .    .  | .
'     +a1n  a2n  a3n  .    ann+ i=n
Public Sub LUSymmetricLinearSolver(A() As Double, B() As Double, x() As Double)
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim ii As Integer, jj As Integer
Dim A1LM1 As Integer, A2LM1 As Integer
Dim bLM1 As Integer, xLM1 As Integer
Dim UpperBound As Integer
Dim tmp As Double, tmp1 As Double
Dim LU() As Double
Dim y() As Double
Dim z() As Double

A1LM1 = LBound(A(), 1) - 1
A2LM1 = LBound(A(), 2) - 1
bLM1 = LBound(B()) - 1
xLM1 = LBound(x()) - 1
UpperBound = UBound(A(), 1) - A1LM1

If UpperBound <> (UBound(A(), 2) - A2LM1) Then
    MsgBox "Parameter array A() is not square!"
End If
If UpperBound <> (UBound(B()) - bLM1) Then
    MsgBox "Number of element of parameter vector b() is not equal to that of A()!"
End If
If UpperBound <> (UBound(x()) - xLM1) Then
    MsgBox "Number of element of parameter vector x() is not equal to that of A()!"
End If

ReDim LU(1 To UpperBound, 1 To UpperBound)
ReDim y(1 To UpperBound)
ReDim z(1 To UpperBound)

'A = transpose(U) * D * U
For j = 1 To UpperBound - 1
    jj = j + A2LM1
'D
    i = j
    tmp = A(i + A1LM1, jj)
    For k = 1 To j - 1
            tmp = tmp - LU(i, k) * LU(k, k) * LU(k, j)
    Next k
    If tmp = 0# Then
        MsgBox "Singular Matrix" & vbNewLine & "Divided by zero", vbCritical
        Exit Sub
    End If
    LU(i, j) = tmp
'U and transpose(U)
    For i = j + 1 To UpperBound
        tmp1 = A(i + A1LM1, jj)
        For k = 1 To j - 1
            tmp1 = tmp1 - LU(i, k) * LU(k, k) * LU(k, j)
        Next k
        tmp1 = tmp1 / tmp
        LU(i, j) = tmp1
        LU(j, i) = tmp1
    Next i
Next j
'D(n, n)
i = UpperBound
j = UpperBound
tmp = A(i + A1LM1, j + A2LM1)
For k = 1 To UpperBound - 1
        tmp = tmp - LU(i, k) * LU(k, k) * LU(k, j)
Next k
If tmp = 0# Then
    MsgBox "Singular Matrix" & vbNewLine & "Divided by zero", vbCritical
    Exit Sub
End If
LU(i, j) = tmp
'transpose(U) * z = b
For i = 1 To UpperBound
    tmp = B(i + bLM1)
    For k = 1 To i - 1
        tmp = tmp - LU(i, k) * z(k)
    Next k
    z(i) = tmp
Next i
'D * y = z
For i = 1 To UpperBound
    y(i) = z(i) / LU(i, i)
Next i
'U * x = y
For i = UpperBound To 1 Step -1
    tmp = y(i)
    For k = i + 1 To UpperBound
        tmp = tmp - LU(i, k) * x(k + xLM1)
    Next k
    x(i + xLM1) = tmp
Next i
End Sub
'consume more time
'solved solution is almost not affected by the vaue of coefficient of matrix
Public Sub HouseholderLinearSolver(A() As Double, B() As Double, x() As Double)
Dim i As Integer, j As Integer, k As Integer
Dim ii As Integer, jj As Integer
Dim A1LM1 As Integer, A2LM1 As Integer
Dim bLM1 As Integer, xLM1 As Integer
Dim UpperBound As Integer
Dim tmp As Double, tmp1 As Double
Dim alpha As Double, ak As Double
Dim h() As Double
Dim d() As Double

A1LM1 = LBound(A(), 1) - 1
A2LM1 = LBound(A(), 2) - 1
bLM1 = LBound(B()) - 1
xLM1 = LBound(x()) - 1
UpperBound = UBound(A(), 1) - A1LM1

If UpperBound <> (UBound(A(), 2) - A2LM1) Then
    MsgBox "Parameter array A() is not square!"
End If
If UpperBound <> (UBound(B()) - bLM1) Then
    MsgBox "Number of element of parameter vector b() is not equal to that of A()!"
End If
If UpperBound <> (UBound(x()) - xLM1) Then
    MsgBox "Number of element of parameter vector x() is not equal to that of A()!"
End If

ReDim h(1 To UpperBound, 1 To UpperBound)
ReDim d(1 To UpperBound)

For i = 1 To UpperBound
    For j = 1 To UpperBound
        h(i, j) = A(i + A1LM1, j + A2LM1)
    Next j
    x(i + xLM1) = B(i + bLM1)
Next i

For j = 1 To UpperBound
    tmp = 0#
    For k = j To UpperBound
        tmp1 = h(k, j)
        tmp = tmp + tmp1 * tmp1
    Next k

'Rank of matrix is less than number of row(number of equation)
    If (tmp = 0#) Then
        MsgBox "matrix is rank deficient"
        Exit Sub
    End If
      
    tmp1 = h(j, j)
    alpha = Sqr(tmp) * Sgn(tmp1)
    ak = 1# / (tmp + alpha * tmp1)
    h(j, j) = tmp1 + alpha
    d(j) = -alpha

    For k = j + 1 To UpperBound
        tmp = 0#
        For i = j To UpperBound
            tmp = tmp + h(i, k) * h(i, j)
        Next i
        tmp = tmp * ak
        For i = j To UpperBound
            h(i, k) = h(i, k) - tmp * h(i, j)
        Next i
    Next k
'Perform update of RHS
    tmp = 0#
    For i = j To UpperBound
        tmp = tmp + x(i + xLM1) * h(i, j)
    Next i
    tmp = tmp * ak
    For i = j To UpperBound
        ii = i + xLM1
        x(ii) = x(ii) - tmp * h(i, j)
    Next i
Next j

'Perform back-substitution
For i = UpperBound To 1 Step -1
    tmp = 0#
    For k = i + 1 To UpperBound
        tmp = tmp + h(i, k) * x(k + xLM1)
    Next k
    ii = i + xLM1
    x(ii) = (x(ii) - tmp) / d(i)
Next i

End Sub
