Attribute VB_Name = "CurveFitting"
Option Explicit
'
'
Public Sub TestFitting()
Dim i As Integer
Dim c1 As Integer, c2 As Integer, c3 As Integer
Const UpperBound As Integer = 100
Const xLM1 As Integer = -23
Const yLM1 As Integer = -29
Dim x() As Double, y() As Double
Dim xtmp() As Double, ytmp() As Double
Dim poly(0 To 4) As Double
Dim polyf(0 To 4) As Double
Dim err(0 To 4) As Double
Const iniLM1 As Integer = -17
Const dcoeLM1 As Integer = -19
Const cLM1 As Integer = -23
Const cfLM1 As Integer = -29
Const eLM1 As Integer = -37
Dim coeInitial(1 + iniLM1 To 3 + iniLM1) As Double
Dim dcoe(1 + dcoeLM1 To 3 + dcoeLM1) As Double
Dim coefficients(1 + cLM1 To 3 + cLM1) As Double
Dim coefficientsf(1 + cfLM1 To 3 + cfLM1) As Double
Dim cerr(1 + eLM1 To 3 + eLM1) As Double
Dim A As Double, B As Double
Dim Af As Double, Bf As Double
Dim message As String

ReDim x(1 + xLM1 To UpperBound + xLM1)
ReDim y(1 + yLM1 To UpperBound + yLM1)
ReDim xtmp(1 + xLM1 To UpperBound + xLM1)
ReDim ytmp(1 + yLM1 To UpperBound + yLM1)
Randomize timer
ThisWorkbook.Worksheets("Fitting").Activate
c1 = -2
c2 = -1
c3 = 0

'LinearFitting
A = 10# * Rnd()
B = 0.1 * Rnd()
message = "LinearFitting" & vbNewLine
message = message & "Actual:" & vbNewLine
message = message & "A: " & A & vbNewLine
message = message & "B: " & B & vbNewLine
message = message & vbNewLine
For i = 1 To UpperBound
    x(i + xLM1) = 100 + 10 * i
    y(i + yLM1) = A + B * (1 + 0.1 * Rnd()) * (x(i + xLM1))
Next i
Call LinearFitting(x(), y(), Af, Bf)
message = message & "Fitting:" & vbNewLine
message = message & "A: " & Af & vbNewLine
message = message & "B: " & Bf & vbNewLine
message = message & vbNewLine
message = message & "Relative err (%):" & vbNewLine
message = message & "A: " & (Af - A) / A * 100 & vbNewLine
message = message & "B: " & (Bf - B) / B * 100 & vbNewLine
MsgBox message
c1 = c1 + 3
c2 = c2 + 3
c3 = c3 + 3
For i = 1 To UpperBound
    Cells(i + 2, c1).Value = x(i + xLM1)
    Cells(i + 2, c2).Value = y(i + yLM1)
    Cells(i + 2, c3).Value = Af + Bf * x(i + xLM1)
Next i

'The value of all elements of array y() should be
'greater than zero!
'ExponentialFitting
A = Rnd()
B = 0.000001 * (Rnd() - 0.5)
message = "ExponentialFitting" & vbNewLine
message = message & "Actual:" & vbNewLine
message = message & "A: " & A & vbNewLine
message = message & "B: " & B & vbNewLine
message = message & vbNewLine
For i = 1 To UpperBound
    x(i + xLM1) = 0.0001 * 1.3 ^ i
    y(i + yLM1) = A * Exp(B * (1 + 0.1 * Rnd()) * (x(i + xLM1)))
Next i
Call ExponentialFitting(x(), y(), Af, Bf)
message = message & "Fitting:" & vbNewLine
message = message & "A: " & Af & vbNewLine
message = message & "B: " & Bf & vbNewLine
message = message & vbNewLine
message = message & "Relative err (%):" & vbNewLine
message = message & "A: " & (Af - A) / A * 100 & vbNewLine
message = message & "B: " & (Bf - B) / B * 100 & vbNewLine
MsgBox message
c1 = c1 + 3
c2 = c2 + 3
c3 = c3 + 3
For i = 1 To UpperBound
    Cells(i + 2, c1).Value = x(i + xLM1)
    Cells(i + 2, c2).Value = y(i + yLM1)
    Cells(i + 2, c3).Value = Af * Exp(Bf * x(i + xLM1))
Next i

'The value of all elements of array y() should be
'greater than zero!
'ExponentialFitting1
A = Rnd()
B = 0.000001 * (Rnd() - 0.5)
message = "ExponentialFitting1" & vbNewLine
message = message & "Actual:" & vbNewLine
message = message & "A: " & A & vbNewLine
message = message & "B: " & B & vbNewLine
message = message & vbNewLine
For i = 1 To UpperBound
    x(i + xLM1) = 0.0001 * 1.3 ^ i
    y(i + yLM1) = A * Exp(B * (1 + 0.1 * Rnd()) * (x(i + xLM1)))
Next i
Call ExponentialFitting1(x(), y(), Af, Bf)
message = message & "Fitting:" & vbNewLine
message = message & "A: " & Af & vbNewLine
message = message & "B: " & Bf & vbNewLine
message = message & vbNewLine
message = message & "Relative err (%):" & vbNewLine
message = message & "A: " & (Af - A) / A * 100 & vbNewLine
message = message & "B: " & (Bf - B) / B * 100 & vbNewLine
MsgBox message
c1 = c1 + 3
c2 = c2 + 3
c3 = c3 + 3
For i = 1 To UpperBound
    Cells(i + 2, c1).Value = x(i + xLM1)
    Cells(i + 2, c2).Value = y(i + yLM1)
    Cells(i + 2, c3).Value = Af * Exp(Bf * x(i + xLM1))
Next i

'The value of all elements of array x() should be
'greater than zero!
'LogarithmicFitting
A = 100# * Rnd()
B = 0.000001 * Rnd()
message = "LogarithmicFitting" & vbNewLine
message = message & "Actual:" & vbNewLine
message = message & "A: " & A & vbNewLine
message = message & "B: " & B & vbNewLine
message = message & vbNewLine
For i = 1 To UpperBound
    x(i + xLM1) = 1# * 1.1 ^ i
    y(i + yLM1) = A + B * Log((1 + 0.1 * Rnd()) * (x(i + xLM1)))
Next i
Call LogarithmicFitting(x(), y(), Af, Bf)
message = message & "Fitting:" & vbNewLine
message = message & "A: " & Af & vbNewLine
message = message & "B: " & Bf & vbNewLine
message = message & vbNewLine
message = message & "Relative err (%):" & vbNewLine
message = message & "A: " & (Af - A) / A * 100 & vbNewLine
message = message & "B: " & (Bf - B) / B * 100 & vbNewLine
MsgBox message
c1 = c1 + 3
c2 = c2 + 3
c3 = c3 + 3
For i = 1 To UpperBound
    Cells(i + 2, c1).Value = x(i + xLM1)
    Cells(i + 2, c2).Value = y(i + yLM1)
    Cells(i + 2, c3).Value = Af + Bf * Log(x(i + xLM1))
Next i

'The value of all elements of array x() and y()should be
'greater than zero!
'PowerLawFitting
A = 200# * Rnd()
B = 0.000001 * Rnd()
message = "PowerLawFitting" & vbNewLine
message = message & "Actual:" & vbNewLine
message = message & "A: " & A & vbNewLine
message = message & "B: " & B & vbNewLine
message = message & vbNewLine
For i = 1 To UpperBound
    x(i + xLM1) = 100 + 10 * i
    y(i + yLM1) = A * (1 + 0.1 * Rnd()) * (x(i + xLM1)) ^ B
Next i
Call PowerLawFitting(x(), y(), Af, Bf)
message = message & "Fitting:" & vbNewLine
message = message & "A: " & Af & vbNewLine
message = message & "B: " & Bf & vbNewLine
message = message & vbNewLine
message = message & "Relative err (%):" & vbNewLine
message = message & "A: " & (Af - A) / A * 100 & vbNewLine
message = message & "B: " & (Bf - B) / B * 100 & vbNewLine
MsgBox message
c1 = c1 + 3
c2 = c2 + 3
c3 = c3 + 3
For i = 1 To UpperBound
    Cells(i + 2, c1).Value = x(i + xLM1)
    Cells(i + 2, c2).Value = y(i + yLM1)
    Cells(i + 2, c3).Value = Af * x(i + xLM1) ^ Bf
Next i

'PolynomialFitting
poly(0) = Rnd()
poly(1) = 0.01 * Rnd()
poly(2) = 0.0001 * Rnd()
poly(3) = 0.000001 * Rnd()
poly(4) = 0.00000001 * Rnd()
message = "PolynomialFitting" & vbNewLine
message = message & "Actual:" & vbNewLine
message = message & ShowPoly(poly())
message = message & vbNewLine
For i = 1 To UpperBound
    x(i + xLM1) = i
    y(i + yLM1) = PolynomialValue(poly(), (1 + 0.1 * Rnd()) * x(i + xLM1))
Next i
Call PolynomialFitting(x(), y(), polyf())
For i = LBound(poly()) To UBound(poly())
    err(i) = (polyf(i) - poly(i)) / poly(i) * 100
Next i
message = message & "Fitting:" & vbNewLine
message = message & ShowPoly(polyf())
message = message & vbNewLine
message = message & "Relative err (%):" & vbNewLine
message = message & ShowPoly(err())
MsgBox message
c1 = c1 + 3
c2 = c2 + 3
c3 = c3 + 3
For i = 1 To UpperBound
    Cells(i + 2, c1).Value = x(i + xLM1)
    Cells(i + 2, c2).Value = y(i + yLM1)
    Cells(i + 2, c3).Value = PolynomialValue(polyf(), x(i + xLM1))
Next i

'NonLinearFitting
coeInitial(1 + iniLM1) = Rnd()
coeInitial(2 + iniLM1) = Rnd()
coeInitial(3 + iniLM1) = Rnd()
dcoe(1 + dcoeLM1) = 0.001
dcoe(2 + dcoeLM1) = 0.001
dcoe(3 + dcoeLM1) = 0.001
coefficients(1 + cLM1) = -0.02
coefficients(2 + cLM1) = 1.86
coefficients(3 + cLM1) = 3.16
message = "NonLinearFitting" & vbNewLine
message = message & "Actual:" & vbNewLine
message = message & ShowPoly(coefficients())
message = message & vbNewLine
For i = 1 To UpperBound
    x(i + xLM1) = i
    xtmp(i + xLM1) = (1 + 0.1 * Rnd()) * x(i + xLM1)
Next i
Call NonlinearFunction(coefficients(), xtmp(), y())
i = NonLinearFitting(x(), y(), _
                     coeInitial(), _
                     dcoe(), _
                     coefficientsf(), _
                     500)
message = message & "Iteration: " & i & vbNewLine
message = message & vbNewLine

Call NonlinearFunction(coefficients(), x(), ytmp())
For i = 1 To 3
    cerr(i + eLM1) = (coefficientsf(i + cfLM1) - coefficients(i + cLM1)) / _
                     coefficients(i + cLM1) * 100
Next i
message = message & "Fitting:" & vbNewLine
message = message & ShowPoly(coefficientsf())
message = message & vbNewLine
message = message & "Relative err (%):" & vbNewLine
message = message & ShowPoly(cerr())
MsgBox message
c1 = c1 + 3
c2 = c2 + 3
c3 = c3 + 3
For i = 1 To UpperBound
    Cells(i + 2, c1).Value = x(i + xLM1)
    Cells(i + 2, c2).Value = y(i + yLM1)
    Cells(i + 2, c3).Value = ytmp(i + yLM1)
Next i

End Sub
'
'f = A + B * x
Public Sub LinearFitting(x() As Double, y() As Double, _
                         A As Double, B As Double)
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer
Dim xSum As Double, ySum As Double
Dim xxSum As Double, xySum As Double
Dim xtmp As Double, ytmp As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Sub
End If

xSum = 0#
ySum = 0#
xxSum = 0#
xySum = 0#
For i = 1 To UpperBound
    xtmp = x(i + xLM1)
    ytmp = y(i + yLM1)
    xSum = xSum + xtmp
    ySum = ySum + ytmp
    xxSum = xxSum + xtmp * xtmp
    xySum = xySum + xtmp * ytmp
Next i

xtmp = 1# / (UpperBound * xxSum - xSum * xSum)
A = (ySum * xxSum - xSum * xySum) * xtmp
B = (UpperBound * xySum - xSum * ySum) * xtmp

End Sub
'The value of all elements of array y() should be
'greater than zero!
'f = A * exp(B * x)
'f = exp(ln(A) + B * x)
'ln(f) = ln(A) + B * x
Public Sub ExponentialFitting(x() As Double, y() As Double, _
                              A As Double, B As Double)
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer
Dim xSum As Double, ySum As Double
Dim xxSum As Double, xySum As Double
Dim xtmp As Double, ytmp As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Sub
End If

xSum = 0#
ySum = 0#
xxSum = 0#
xySum = 0#
For i = 1 To UpperBound
    xtmp = x(i + xLM1)
    ytmp = Log(y(i + yLM1))
    xSum = xSum + xtmp
    ySum = ySum + ytmp
    xxSum = xxSum + xtmp * xtmp
    xySum = xySum + xtmp * ytmp
Next i

xtmp = 1# / (UpperBound * xxSum - xSum * xSum)
A = (ySum * xxSum - xSum * xySum) * xtmp
B = (UpperBound * xySum - xSum * ySum) * xtmp
A = Exp(A)
End Sub
'The value of all elements of array y() should be
'greater than zero!
'f = A * exp(B * x)
'f = exp(ln(A) + B * x)
'ln(f) = ln(A) + B * x
'minimize the function:
'sum(y * (ln(y) - (A + B * x)) ^ 2)
Public Sub ExponentialFitting1(x() As Double, y() As Double, _
                              A As Double, B As Double)
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer
Dim ySum As Double, xySum As Double, xxySum As Double
Dim yLnYSum As Double, xyLnYSum As Double
Dim xtmp As Double, ytmp As Double, LnYTmp As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Sub
End If

ySum = 0#
xySum = 0#
xxySum = 0#
yLnYSum = 0#
xyLnYSum = 0#
For i = 1 To UpperBound
    xtmp = x(i + xLM1)
    ytmp = y(i + yLM1)
    LnYTmp = Log(ytmp)
    ySum = ySum + ytmp
    xySum = xySum + xtmp * ytmp
    xxySum = xxySum + xtmp * xtmp * ytmp
    yLnYSum = yLnYSum + ytmp * LnYTmp
    xyLnYSum = xyLnYSum + xtmp * ytmp * LnYTmp
Next i

xtmp = 1# / (ySum * xxySum - xySum * xySum)
A = (xxySum * yLnYSum - xySum * xyLnYSum) * xtmp
B = (ySum * xyLnYSum - xySum * yLnYSum) * xtmp
A = Exp(A)
End Sub
'The value of all elements of array x() should be
'greater than zero!
'f = A + B * ln(x)
Public Sub LogarithmicFitting(x() As Double, y() As Double, _
                              A As Double, B As Double)
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer
Dim xSum As Double, ySum As Double
Dim xxSum As Double, xySum As Double
Dim xtmp As Double, ytmp As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Sub
End If

xSum = 0#
ySum = 0#
xxSum = 0#
xySum = 0#
For i = 1 To UpperBound
    xtmp = Log(x(i + xLM1))
    ytmp = y(i + yLM1)
    xSum = xSum + xtmp
    ySum = ySum + ytmp
    xxSum = xxSum + xtmp * xtmp
    xySum = xySum + xtmp * ytmp
Next i

B = (UpperBound * xySum - xSum * ySum) / _
    (UpperBound * xxSum - xSum * xSum)
A = (ySum - B * xSum) / UpperBound
End Sub
'The value of all elements of array x() and y()should be
'greater than zero!
'f = A * x ^ B
Public Sub PowerLawFitting(x() As Double, y() As Double, _
                           A As Double, B As Double)
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer
Dim xSum As Double, ySum As Double
Dim xxSum As Double, xySum As Double
Dim xtmp As Double, ytmp As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Sub
End If

xSum = 0#
ySum = 0#
xxSum = 0#
xySum = 0#
For i = 1 To UpperBound
    xtmp = Log(x(i + xLM1))
    ytmp = Log(y(i + yLM1))
    xSum = xSum + xtmp
    ySum = ySum + ytmp
    xxSum = xxSum + xtmp * xtmp
    xySum = xySum + xtmp * ytmp
Next i

B = (UpperBound * xySum - xSum * ySum) / _
    (UpperBound * xxSum - xSum * xSum)
A = (ySum - B * xSum) / UpperBound
A = Exp(A)
End Sub
'
'f = poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub PolynomialFitting(x() As Double, y() As Double, _
                             polynomial() As Double)
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim PolyUpper As Integer
Dim PolyLower As Integer
Dim i As Integer, j As Integer, k As Integer
Dim xtmp As Double, xtmp1 As Double
Dim vector() As Double
Dim matrix() As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Sub
End If

PolyLower = LBound(polynomial)
If PolyLower < 0 Then
    MsgBox "The lower bound of parameter polynomial() is less than 0!"
    Exit Sub
End If
PolyUpper = UBound(polynomial)

ReDim vector(PolyLower To PolyUpper)
ReDim matrix(PolyLower To PolyUpper, PolyLower To PolyUpper)
For i = PolyLower To PolyUpper
    vector(i) = 0#
    For j = PolyLower To PolyUpper
        matrix(i, j) = 0#
    Next j
Next i

For i = PolyLower To PolyUpper
    For k = 1 To UpperBound
        xtmp = x(k + xLM1)
        xtmp1 = power(xtmp, i)
        vector(i) = vector(i) + y(k + yLM1) * xtmp1
        matrix(i, i) = matrix(i, i) + xtmp1 * xtmp1
        For j = i + 1 To PolyUpper
            matrix(i, j) = matrix(i, j) + power(xtmp, j) * xtmp1
        Next j
    Next k
    For j = i + 1 To PolyUpper
        matrix(j, i) = matrix(i, j)
    Next j
Next i

Call LUSymmetricLinearSolver(matrix(), vector(), polynomial())

End Sub
'
'
Public Function power(x As Double, n As Integer) As Double
Dim i As Integer
If n >= 0 Then
    If n < 40 Then
        power = 1#
        For i = 1 To n
            power = power * x
        Next i
    Else
        power = x ^ n
    End If
Else
    If n > -40 Then
        power = 1#
        x = 1# / x
        n = -n
        For i = 1 To n
            power = power * x
        Next i
    Else
        power = x ^ n
    End If
End If
End Function
'Curve fitting for a non-linear function is basically
'an optimization problem for minimizing the fitting err,
'you can use any method in the Optimization module to
'solve this problem, and replace the Fittness fucntion
'with the function for calculating fitting err.
'Here the NewtonOptimization is adopted.
'f() = NonlinearFunction()
Public Function NonLinearFitting(x() As Double, _
                    y() As Double, _
                    coeInitial() As Double, _
                    dcoe() As Double, _
                    coefficients() As Double, _
                    MaxIteration As Integer) As Integer
Dim i As Integer, j As Integer
Dim ii As Integer, jj As Integer
Dim IterationCounter As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim iniLM1 As Integer, dcoeLM1 As Integer, coeLM1 As Integer
Dim coeUpperBound As Integer
Dim tmp1 As Double, tmp2 As Double
Dim dci As Double, dcj As Double
Dim ciPdci As Double, ciMdci As Double
Dim ec As Double
Dim ePP As Double, eMM As Double
Dim ePM As Double, eMP As Double
Dim deNorm As Double
Dim tolerence As Double
Dim coeTmp() As Double
Dim Mde() As Double
Dim cStep() As Double
Dim dcInv() As Double
Dim dc2Inv() As Double
Dim Jacobian() As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Function
End If

iniLM1 = LBound(coeInitial) - 1
dcoeLM1 = LBound(dcoe) - 1
coeLM1 = LBound(coefficients) - 1
coeUpperBound = UBound(coeInitial) - LBound(coeInitial) + 1

If coeUpperBound <> (UBound(dcoe) - dcoeLM1) Then
    MsgBox "Number of element of Parameter array coeInitial() " & vbNewLine & _
           "is not equal to that of array dcoe()!"
    Exit Function
End If
If coeUpperBound <> (UBound(coefficients) - coeLM1) Then
    MsgBox "Number of element of Parameter array coeInitial() " & vbNewLine & _
           "is not equal to that of array coefficients()!"
    Exit Function
End If

ReDim coeTmp(1 + iniLM1 To coeUpperBound + iniLM1)
ReDim Mde(1 To coeUpperBound)
ReDim cStep(1 To coeUpperBound)
ReDim dcInv(1 To coeUpperBound)
ReDim dc2Inv(1 To coeUpperBound, 1 To coeUpperBound)
ReDim Jacobian(1 To coeUpperBound, 1 To coeUpperBound)
'initiate
For i = 1 To coeUpperBound
    ii = i + iniLM1
    coeTmp(ii) = coeInitial(ii)
Next i

tolerence = 0.000000001 * Sqr(CDbl(coeUpperBound))
deNorm = 0#
ec = NonlinearErr(coeTmp(), x(), y())
For i = 1 To coeUpperBound
    dci = dcoe(i + dcoeLM1)
'initiate dcInv(i) and dc2Inv(i, j)
    dcInv(i) = 1# / (2# * dci)
    dc2Inv(i, i) = 1# / (dci * dci)
    For j = i + 1 To coeUpperBound
        tmp1 = 1# / (2# * dci * 2# * dcoe(j + dcoeLM1))
        dc2Inv(i, j) = tmp1
        dc2Inv(j, i) = tmp1
    Next j
    ii = i + iniLM1
    tmp1 = coeTmp(ii)
    coeTmp(ii) = tmp1 + dci
    ePP = NonlinearErr(coeTmp(), x(), y())
    coeTmp(ii) = tmp1 - dci
    eMM = NonlinearErr(coeTmp(), x(), y())
    coeTmp(ii) = tmp1
'-de and Norm(central difference)
    tmp2 = (ePP - eMM) * dcInv(i)
    Mde(i) = -tmp2
    deNorm = deNorm + tmp2 * tmp2
'diagonal element of Jacobian matrix
    Jacobian(i, i) = (ePP - 2# * ec + eMM) * dc2Inv(i, i)
Next i
deNorm = Sqr(deNorm)
IterationCounter = 0
Do While (deNorm > tolerence)
    IterationCounter = IterationCounter + 1
    If IterationCounter > MaxIteration Then Exit Do
'off-diagonal element of Jacobian matrix
    For i = 1 To coeUpperBound - 1
        dci = dcoe(i + dcoeLM1)
        ii = i + iniLM1
        tmp1 = coeTmp(ii)
        ciPdci = tmp1 + dci
        ciMdci = tmp1 - dci
        For j = i + 1 To coeUpperBound
            dcj = dcoe(j + dcoeLM1)
            jj = j + iniLM1
            tmp2 = coeTmp(jj)

            coeTmp(ii) = ciPdci
            coeTmp(jj) = tmp2 + dcj
            ePP = NonlinearErr(coeTmp(), x(), y())
            coeTmp(ii) = ciMdci
            eMP = NonlinearErr(coeTmp(), x(), y())
            coeTmp(jj) = tmp2 - dcj
            eMM = NonlinearErr(coeTmp(), x(), y())
            coeTmp(ii) = ciPdci
            ePM = NonlinearErr(coeTmp(), x(), y())
            
            coeTmp(jj) = tmp2
            tmp2 = (ePP - eMP - ePM + eMM) * dc2Inv(i, j)
            Jacobian(i, j) = tmp2
            Jacobian(j, i) = tmp2
        Next j
        coeTmp(ii) = tmp1
    Next i
'call LU decomposition to solve
    Call LUSymmetricLinearSolver(Jacobian(), Mde(), cStep())
'update x
    For i = 1 To coeUpperBound
        ii = i + iniLM1
        coeTmp(ii) = coeTmp(ii) + cStep(i)
    Next i
'Norm and diagonal element
    deNorm = 0#
    ec = NonlinearErr(coeTmp(), x(), y())
    For i = 1 To coeUpperBound
        dci = dcoe(i + dcoeLM1)
        ii = i + iniLM1
        tmp1 = coeTmp(ii)
        coeTmp(ii) = tmp1 + dci
        ePP = NonlinearErr(coeTmp(), x(), y())
        coeTmp(ii) = tmp1 - dci
        eMM = NonlinearErr(coeTmp(), x(), y())
        coeTmp(ii) = tmp1
'-de and Norm(central difference)
        tmp2 = (ePP - eMM) * dcInv(i)
        Mde(i) = -tmp2
        deNorm = deNorm + tmp2 * tmp2
'diagonal element of Jacobian matrix
        Jacobian(i, i) = (ePP - 2# * ec + eMM) * dc2Inv(i, i)
    Next i
    deNorm = Sqr(deNorm)
Loop

NonLinearFitting = IterationCounter
For i = 1 To coeUpperBound
    coefficients(i + coeLM1) = coeTmp(i + iniLM1)
Next i

End Function
'
'subroutine for calculating fitted value of non-linear function
Public Sub NonlinearFunction(coefficients() As Double, _
                             x() As Double, _
                             FittingValue() As Double)
Dim i As Integer
Dim xLM1 As Integer, fittingLM1 As Integer, cLM1 As Integer
Dim UpperBound As Integer
Dim xtmp As Double
'¡õ¡õ¡õDeclare coefficients you need here¡õ¡õ¡õ
Dim c1 As Double
Dim c2 As Double
Dim c3 As Double

xLM1 = LBound(x) - 1
fittingLM1 = LBound(FittingValue) - 1
UpperBound = UBound(x) - xLM1

cLM1 = LBound(coefficients) - 1
'¡õ¡õ¡õSet coefficients here¡õ¡õ¡õ
c1 = coefficients(1 + cLM1)
c2 = coefficients(2 + cLM1)
c3 = coefficients(3 + cLM1)

For i = 1 To UpperBound
    xtmp = x(i + xLM1)
'¡õ¡õ¡õSet non-linear function here¡õ¡õ¡õ
    FittingValue(i + fittingLM1) = c1 * xtmp * xtmp + _
                                   c2 * xtmp + _
                                   c3
Next i

End Sub
'
'error for non-linear fitting
Private Function NonlinearErr(coefficients() As Double, _
                             x() As Double, _
                             y() As Double) As Double
Dim i As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim UpperBound As Integer
Dim tmp As Double
Dim FittingValue() As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
ReDim FittingValue(1 + yLM1 To UpperBound + yLM1)

Call NonlinearFunction(coefficients(), x(), FittingValue())

NonlinearErr = 0#
For i = 1 To UpperBound
    tmp = FittingValue(i + yLM1) - y(i + yLM1)
    NonlinearErr = NonlinearErr + tmp * tmp
Next i

End Function
