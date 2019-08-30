Attribute VB_Name = "Root"
Option Explicit
'
'
Public Sub TestRoot()
Const xLM1 As Integer = -11
Const dLM1 As Integer = -23
Const solLM1 As Integer = -37
Dim i As Integer, iteration As Integer
Dim n As Integer
Const dx As Double = 0.0000001
Dim xLower As Double, xUpper As Double
Dim solNewton As Double
Dim solHalley As Double
Dim solSchroder As Double
Dim solSteffenson As Double
Dim solBiSection As Double
Dim solFalsePosition As Double
Dim solBrent As Double
Dim TimerStart As Double, TimerEnd As Double
Dim tmp As Double
Dim xInit(1 + xLM1 To 2 + xLM1) As Double
Dim d(1 + dLM1 To 2 + dLM1) As Double
Dim sol(1 + solLM1 To 2 + solLM1) As Double
Dim message As String

n = 1000
xLower = -0.7
xUpper = 0.5
message = ""
'Newton
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = Newton(xLower, dx, solNewton, 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "Newton: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & solNewton & vbNewLine
message = message & "residue: " & RootFunction(solNewton) & vbNewLine
message = message & vbNewLine
'Halley
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = Halley(xLower, dx, solHalley, 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "Halley: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & solHalley & vbNewLine
message = message & "residue: " & RootFunction(solHalley) & vbNewLine
message = message & vbNewLine
'Schroder
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = Schroder(xLower, dx, solSchroder, 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "Schroder: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & solSchroder & vbNewLine
message = message & "residue: " & RootFunction(solSchroder) & vbNewLine
message = message & vbNewLine
'Steffenson
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = Steffenson(xLower, dx, solSteffenson, 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "Steffenson: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & solSteffenson & vbNewLine
message = message & "residue: " & RootFunction(solSteffenson) & vbNewLine
message = message & vbNewLine
'BiSection
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = BiSection(xLower, xUpper, solBiSection, 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "BiSection: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & solBiSection & vbNewLine
message = message & "residue: " & RootFunction(solBiSection) & vbNewLine
message = message & vbNewLine
'FalsePosition
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = FalsePosition(xLower, xUpper, solFalsePosition, 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "FalsePosition: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & solFalsePosition & vbNewLine
message = message & "residue: " & RootFunction(solFalsePosition) & vbNewLine
message = message & vbNewLine
'Brent
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = Brent(xLower, xUpper, solBrent, 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "Brent: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & solBrent & vbNewLine
message = message & "residue: " & RootFunction(solBrent) & vbNewLine
message = message & vbNewLine

MsgBox message


n = 500
For i = 1 To 2
    xInit(i + xLM1) = 1#
    d(i + dLM1) = 0.0000001
Next i
message = ""
'SimultaneousNewton
tmp = 0#
For i = 1 To n
    TimerStart = timer
    iteration = SimultaneousNewton(xInit(), d(), sol(), 100)
    TimerEnd = timer
    tmp = tmp + (TimerEnd - TimerStart)
Next i
message = message & "SimultaneousNewton: " & tmp & " seconds" & vbNewLine
message = message & "iteration: " & iteration & vbNewLine
message = message & "solution: " & vbNewLine & ShowPoly(sol) & vbNewLine
message = message & "residue: " & vbNewLine
message = message & RootFunction1(sol) & vbNewLine
message = message & RootFunction2(sol) & vbNewLine
MsgBox message
End Sub
'
'xNew = xOld - f(xOld) / f'(xOld)
Public Function Newton(xInitial As Double, dx As Double, _
                       solution As Double, MaxIteration As Integer) As Integer
Dim IterationCounter As Integer
Dim xOld As Double, xNew As Double
Dim f As Double, fP As Double, fM As Double
Const tolerence As Double = 0.000000001

If MaxIteration < 1 Then
    MsgBox "Parameter MaxIteration should be greater than or equal to 1."
    Exit Function
End If

xOld = xInitial
f = RootFunction(xOld)
If Abs(f) < tolerence Then
    solution = xOld
    Newton = 0
    Exit Function
End If
For IterationCounter = 1 To MaxIteration
    fP = RootFunction(xOld + dx)
    fM = RootFunction(xOld - dx)
    xNew = xOld - f * 2# * dx / (fP - fM)
    f = RootFunction(xNew)
    If Abs(f) < tolerence Then Exit For
    xOld = xNew
Next IterationCounter

solution = xNew
Newton = IterationCounter
End Function
'
'xNew = xOld - 2 * f(xOld) * f'(xOld) / (2 * f'(xOld) ^ 2 - f(xOld) * f''(xOld))
Public Function Halley(xInitial As Double, dx As Double, _
                       solution As Double, MaxIteration As Integer) As Integer
Dim IterationCounter As Integer
Dim xOld As Double, xNew As Double
Dim f As Double, fP As Double, fM As Double
Dim fPrime As Double, fPrimePrime As Double
Dim Twodxinv As Double, dxSquareinv As Double
Const tolerence As Double = 0.000000001

If MaxIteration < 1 Then
    MsgBox "Parameter MaxIteration should be greater than or equal to 1."
    Exit Function
End If

xOld = xInitial
f = RootFunction(xOld)
If Abs(f) < tolerence Then
    solution = xOld
    Halley = 0
    Exit Function
End If
Twodxinv = 0.5 / dx
dxSquareinv = 1# / dx / dx
For IterationCounter = 1 To MaxIteration
    fP = RootFunction(xOld + dx)
    fM = RootFunction(xOld - dx)
    fPrime = (fP - fM) * Twodxinv
    fPrimePrime = (fP - 2# * f + fM) * dxSquareinv
    fP = 2# * fPrime * fPrime - f * fPrimePrime
    If fP <> 0 Then
        xNew = xOld - 2# * f * fPrime / fP
    Else
        xNew = xOld - f / fPrime
    End If
    f = RootFunction(xNew)
    If Abs(f) < tolerence Then Exit For
    xOld = xNew
Next IterationCounter

solution = xNew
Halley = IterationCounter
End Function
'
'xNew = xOld - f(xOld) * f'(xOld) / (f'(xOld) ^ 2 - f(xOld) * f''(xOld))
Public Function Schroder(xInitial As Double, dx As Double, _
                         solution As Double, MaxIteration As Integer) As Integer
Dim IterationCounter As Integer
Dim xOld As Double, xNew As Double
Dim f As Double, fP As Double, fM As Double
Dim fPrime As Double, fPrimePrime As Double
Dim Twodxinv As Double, dxSquareinv As Double
Const tolerence As Double = 0.000000001

If MaxIteration < 1 Then
    MsgBox "Parameter MaxIteration should be greater than or equal to 1."
    Exit Function
End If

xOld = xInitial
f = RootFunction(xOld)
If Abs(f) < tolerence Then
    solution = xOld
    Schroder = 0
    Exit Function
End If
Twodxinv = 0.5 / dx
dxSquareinv = 1# / dx / dx
For IterationCounter = 1 To MaxIteration
    fP = RootFunction(xOld + dx)
    fM = RootFunction(xOld - dx)
    fPrime = (fP - fM) * Twodxinv
    fPrimePrime = (fP - 2# * f + fM) * dxSquareinv
    fP = fPrime * fPrime - f * fPrimePrime
    If fP <> 0 Then
        xNew = xOld - f * fPrime / fP
    Else
        xNew = xOld - f / fPrime
    End If
    f = RootFunction(xNew)
    If Abs(f) < tolerence Then Exit For
    xOld = xNew
Next IterationCounter

solution = xNew
Schroder = IterationCounter
End Function
'
'
Public Function Steffenson(xInitial As Double, dx As Double, _
                           solution As Double, MaxIteration As Integer) As Integer
Dim IterationCounter As Integer
Dim x As Double, x1 As Double, xNew As Double
Dim f As Double, fP As Double, fM As Double
Const tolerence As Double = 0.000000001

If MaxIteration < 1 Then
    MsgBox "Parameter MaxIteration should be greater than or equal to 1."
    Exit Function
End If

x = xInitial
f = RootFunction(x)
If Abs(f) < tolerence Then
    solution = x
    Steffenson = 0
    Exit Function
End If
For IterationCounter = 1 To MaxIteration
    fP = RootFunction(x + dx)
    fM = RootFunction(x - dx)
    xNew = x - f * 2# * dx / (fP - fM)
    f = RootFunction(xNew)
    If Abs(f) < tolerence Then Exit For
    If (IterationCounter >= 3) Then
        fP = xNew - x1
        fM = 2# * (xNew - 2# * x + x1)
        If (fM <> 0) Then
            xNew = x - fP * fP / fM
            f = RootFunction(xNew)
            If Abs(f) < tolerence Then Exit For
            x = xNew
        End If
    End If
    x1 = x
    x = xNew
Next IterationCounter

solution = xNew
Steffenson = IterationCounter

End Function
'
'
Public Function BiSection(x0 As Double, x1 As Double, _
                          solution As Double, MaxIteration As Integer) As Integer
Dim IterationCounter As Integer
Dim xLower As Double, xUpper As Double, x As Double
Dim fL As Double, fU As Double, f As Double
Const tolerence As Double = 0.000000001

If MaxIteration < 1 Then
    MsgBox "Parameter MaxIteration should be greater than or equal to 1."
    Exit Function
End If

If x0 = x1 Then
    MsgBox "The value of parameter x0 should differ from that of x1."
    Exit Function
End If

If x0 < x1 Then
    xLower = x0
    xUpper = x1
Else
    xLower = x1
    xUpper = x0
End If
fL = RootFunction(xLower)
If Abs(fL) < tolerence Then
    solution = xLower
    BiSection = 0
    Exit Function
End If
fU = RootFunction(xUpper)
If Abs(fU) < tolerence Then
    solution = xUpper
    BiSection = 0
    Exit Function
End If

If fL * fU > 0 Then
    MsgBox "The sign of function value at x0 should differ from that at x1."
    Exit Function
End If

For IterationCounter = 1 To MaxIteration
    x = 0.5 * (xLower + xUpper)
    f = RootFunction(x)
    If Abs(f) < tolerence Then Exit For
    If f * fL > 0 Then
        xLower = x
        fL = f
    Else
        xUpper = x
        fU = f
    End If
Next IterationCounter

solution = x
BiSection = IterationCounter
End Function
'
'
Public Function FalsePosition(x0 As Double, x1 As Double, _
                              solution As Double, MaxIteration As Integer) As Integer
Dim IterationCounter As Integer
Dim xLower As Double, xUpper As Double, x As Double
Dim fL As Double, fU As Double, f As Double
Const tolerence As Double = 0.000000001

If MaxIteration < 1 Then
    MsgBox "Parameter MaxIteration should be greater than or equal to 1."
    Exit Function
End If

If x0 = x1 Then
    MsgBox "The value of parameter x0 should differ from that of x1."
    Exit Function
End If

If x0 < x1 Then
    xLower = x0
    xUpper = x1
Else
    xLower = x1
    xUpper = x0
End If
fL = RootFunction(xLower)
If Abs(fL) < tolerence Then
    solution = xLower
    FalsePosition = 0
    Exit Function
End If
fU = RootFunction(xUpper)
If Abs(fU) < tolerence Then
    solution = xUpper
    FalsePosition = 0
    Exit Function
End If

If fL * fU > 0 Then
    MsgBox "The sign of function value at x0 should differ from that at x1."
    Exit Function
End If

For IterationCounter = 1 To MaxIteration
    x = xLower - fL * (xUpper - xLower) / (fU - fL)
    f = RootFunction(x)
    If Abs(f) < tolerence Then Exit For
    If f * fL > 0 Then
        xLower = x
        fL = f
    Else
        xUpper = x
        fU = f
    End If
Next IterationCounter

solution = x
FalsePosition = IterationCounter
End Function
'
'
Public Function Brent(x0 As Double, x1 As Double, _
                      solution As Double, MaxIteration As Integer) As Integer
Dim IterationCounter As Integer
Dim xLower As Double, xUpper As Double, x As Double
Dim fL As Double, fU As Double, f As Double
Dim xNew As Double, fNew As Double
Dim fMfLinv As Double, fMfUinv As Double, fLMfUinv As Double
Const tolerence As Double = 0.000000001

If MaxIteration < 1 Then
    MsgBox "Parameter MaxIteration should be greater than or equal to 1."
    Exit Function
End If

If x0 = x1 Then
    MsgBox "The value of parameter x0 should differ from that of x1."
    Exit Function
End If

If x0 < x1 Then
    xLower = x0
    xUpper = x1
Else
    xLower = x1
    xUpper = x0
End If
fL = RootFunction(xLower)
If Abs(fL) < tolerence Then
    solution = xLower
    Brent = 0
    Exit Function
End If
fU = RootFunction(xUpper)
If Abs(fU) < tolerence Then
    solution = xUpper
    Brent = 0
    Exit Function
End If

If fL * fU > 0 Then
    MsgBox "The sign of function value at x0 should differ from that at x1."
    Exit Function
End If

'False position for 1st iteration
x = xLower - fL * (xUpper - xLower) / (fU - fL)
f = RootFunction(x)
If Abs(f) < tolerence Then
    solution = x
    Brent = 1
    Exit Function
End If
    
For IterationCounter = 2 To MaxIteration
    fMfLinv = f - fL
    fMfUinv = f - fU
    fLMfUinv = fL - fU
    If (fMfLinv <> 0) Or (fMfUinv <> 0) Or _
       (fLMfUinv <> 0) Then
        fMfLinv = 1# / fMfLinv
        fMfUinv = 1# / fMfUinv
        fLMfUinv = 1# / fLMfUinv
        xNew = x * fL * fU * fMfLinv * fMfUinv + _
               xLower * f * fU * (-fMfLinv) * fLMfUinv + _
               xUpper * f * fL * (-fMfUinv) * (-fLMfUinv)
        fNew = RootFunction(xNew)
        If Abs(fNew) < tolerence Then Exit For
'xNew replaces the closest one of xLower, x and xUpper
        If xNew >= x Then
            If Abs(xNew - x) < Abs(xNew - xUpper) Then
                x = xNew
                f = fNew
            Else
                xUpper = xNew
                fU = fNew
            End If
        Else
            If Abs(xNew - x) < Abs(xNew - xLower) Then
                x = xNew
                f = fNew
            Else
                xLower = xNew
                fL = fNew
            End If
        End If
'False position
'If false position is executed twice without interruption,
'then false position will be executed repeatedly.
    Else
        x = xLower - fL * (xUpper - xLower) / (fU - fL)
        f = RootFunction(x)
        If Abs(f) < tolerence Then
            xNew = x
            Exit For
        End If
    End If
Next IterationCounter

solution = xNew
Brent = IterationCounter

End Function
'
'
Public Function SimultaneousNewton(xInitial() As Double, _
                                   dx() As Double, _
                                   solution() As Double, _
                                   MaxIteration As Integer) As Integer
Dim i As Integer, j As Integer
Dim ii As Integer, jj As Integer
Dim IterationCounter As Integer
Dim iniLM1 As Integer, dxLM1 As Integer, solLM1 As Integer
Dim UpperBound As Integer
Dim tmp As Double, tmp1 As Double
Dim dxi As Double, dxj As Double
Dim xiPdxi As Double, xiMdxi As Double
Dim fNorm As Double
Dim tolerence As Double
Dim xtmp() As Double
Dim Twodxinv() As Double
Dim df() As Double
Dim xStep() As Double
Dim Mf() As Double, fP() As Double, fM() As Double

iniLM1 = LBound(xInitial) - 1
dxLM1 = LBound(dx) - 1
solLM1 = LBound(solution) - 1
UpperBound = UBound(xInitial) - LBound(xInitial) + 1
'The number of function to be solved is 2 in this example!
If UpperBound <> 2 Then
    MsgBox "The number of function to be solved is 2 in this example!"
    Exit Function
End If
If UpperBound <> (UBound(dx) - dxLM1) Then
    MsgBox "Number of element of Parameter array xInitial() " & vbNewLine & _
           "is not equal to that of array dx()!"
    Exit Function
End If
If UpperBound <> (UBound(solution) - solLM1) Then
    MsgBox "Number of element of Parameter array xInitial() " & vbNewLine & _
           "is not equal to that of array Solution()!"
    Exit Function
End If

ReDim xtmp(1 + iniLM1 To UpperBound + iniLM1)
ReDim Twodxinv(1 To UpperBound)
ReDim df(1 To UpperBound, 1 To UpperBound)
ReDim xStep(1 To UpperBound)
ReDim Mf(1 To UpperBound)
ReDim fP(1 To UpperBound)
ReDim fM(1 To UpperBound)
'initiate
For i = 1 To UpperBound
    ii = i + iniLM1
    xtmp(ii) = xInitial(ii)
    Twodxinv(i) = 1# / (2# * dx(i + dxLM1))
Next i

tolerence = 0.000000001 * Sqr(CDbl(UpperBound))
'Norm and -f vector
fNorm = 0#
tmp = RootFunction1(xtmp())
fNorm = fNorm + tmp * tmp
Mf(1) = -tmp
tmp = RootFunction2(xtmp())
fNorm = fNorm + tmp * tmp
Mf(2) = -tmp
fNorm = Sqr(fNorm)
IterationCounter = 0
Do While (fNorm > tolerence)
    IterationCounter = IterationCounter + 1
    If IterationCounter > MaxIteration Then Exit Do
'column 1 of df matrix (for variable 1)
    tmp = xtmp(1 + iniLM1)
    tmp1 = dx(1 + dxLM1)
    xtmp(1 + iniLM1) = tmp + tmp1
    fP(1) = RootFunction1(xtmp())
    fP(2) = RootFunction2(xtmp())
    xtmp(1 + iniLM1) = tmp - tmp1
    fM(1) = RootFunction1(xtmp())
    fM(2) = RootFunction2(xtmp())
    xtmp(1 + iniLM1) = tmp
    df(1, 1) = (fP(1) - fM(1)) * Twodxinv(1)
    df(2, 1) = (fP(2) - fM(2)) * Twodxinv(2)
'column 2 of df matrix (for variable 2)
    tmp = xtmp(2 + iniLM1)
    tmp1 = dx(2 + dxLM1)
    xtmp(2 + iniLM1) = tmp + tmp1
    fP(1) = RootFunction1(xtmp())
    fP(2) = RootFunction2(xtmp())
    xtmp(2 + iniLM1) = tmp - tmp1
    fM(1) = RootFunction1(xtmp())
    fM(2) = RootFunction2(xtmp())
    xtmp(2 + iniLM1) = tmp
    df(1, 2) = (fP(1) - fM(1)) * Twodxinv(1)
    df(2, 2) = (fP(2) - fM(2)) * Twodxinv(2)
'call LU decomposition to solve
    Call LULinearSolver(df(), Mf(), xStep())
'update x
    For i = 1 To UpperBound
        ii = i + iniLM1
        xtmp(ii) = xtmp(ii) + xStep(i)
    Next i
'Norm and -f vector
    fNorm = 0#
    tmp = RootFunction1(xtmp())
    fNorm = fNorm + tmp * tmp
    Mf(1) = -tmp
    tmp = RootFunction2(xtmp())
    fNorm = fNorm + tmp * tmp
    Mf(2) = -tmp
    fNorm = Sqr(fNorm)
Loop

SimultaneousNewton = IterationCounter
For i = 1 To UpperBound
    solution(i + solLM1) = xtmp(i + iniLM1)
Next i
End Function
'
'
Public Function RootFunction(x As Double) As Double
RootFunction = (Tan(x) - Tan(-x)) / (Exp(x ^ 2) + Exp(-x ^ 2))
'RootFunction = 1# / (1# + x * x) - 0.5
End Function
'
'
Public Function RootFunction1(x() As Double) As Double
Dim l As Integer
Dim x1 As Double, x2 As Double

l = LBound(x) - 1
x1 = x(1 + l)
x2 = x(2 + l)

RootFunction1 = x1 ^ 2 + x2 ^ 2 - 1#
End Function
'
'
Public Function RootFunction2(x() As Double) As Double
Dim l As Integer
Dim x1 As Double, x2 As Double

l = LBound(x) - 1
x1 = x(1 + l)
x2 = x(2 + l)

RootFunction2 = 5# * x1 + 3# * x2 + 2#
End Function
