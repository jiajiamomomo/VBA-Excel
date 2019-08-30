Attribute VB_Name = "Optimization"
Option Explicit
'For unconstrained optimization problem only.
'
'function for local extremum(return the number of iteration
'when the solution converged):
'NewtonOptimization
'SteepestDescent
'ConjugateGradientFR
'ConjugateGradientPR
'
'subroutine for global minimum:
'SimulatedAnnealing
'
'subroutine for global maximum:
'GeneticAlgorithm
'
'name of target function is "Fitness"
'Fitness has parameter: array x() As Double
Public Sub TestOptimization()
Dim TimerStart As Double, TimerEnd As Double
Dim TimerNewton As Double
Dim TimerSteepest As Double
Dim TimerFletcherReeves As Double
Dim TimerPolakRibiere As Double
Dim TimerSimulatedAnnealing As Double
Dim TimerGeneticAlgorithm As Double
Dim xInitial() As Double
Dim dx() As Double
Dim solution() As Double
Const n As Integer = 2
Const InitLM1 As Integer = 11
Const dxLM1 As Integer = 15
Const solLM1 As Integer = 17
Dim i As Integer, j As Integer
Dim message As String

ReDim xInitial(1 + InitLM1 To n + InitLM1)
ReDim dx(1 + dxLM1 To n + dxLM1)
ReDim solution(1 + solLM1 To n + solLM1)
xInitial(1 + InitLM1) = 10#
xInitial(2 + InitLM1) = 10#
dx(1 + dxLM1) = 0.0000001
dx(2 + dxLM1) = 0.0000001

'NewtonOptimization
TimerStart = timer
For j = 1 To 1000
i = NewtonOptimization(xInitial(), dx(), solution(), 500)
Next j
TimerEnd = timer
TimerNewton = TimerEnd - TimerStart
If i > 500 Then
    MsgBox "NewtonOptimization Not converged yet."
Else
    MsgBox "NewtonOptimization Converged after " & i & " steps."
End If
message = "Solution of NewtonOptimization: " & vbTab & TimerNewton & " seconds" & vbNewLine
For i = 1 To n
    message = message & i & vbTab & solution(i + solLM1) & vbNewLine
Next i
MsgBox message

'SteepestDescent
TimerStart = timer
For j = 1 To 1000
i = SteepestDescent(xInitial(), dx(), solution(), 500)
Next j
TimerEnd = timer
TimerSteepest = TimerEnd - TimerStart
If i > 500 Then
    MsgBox "SteepestDescent Not converged yet."
Else
    MsgBox "SteepestDescent Converged after " & i & " steps."
End If
message = "Solution of SteepestDescent: " & vbTab & TimerSteepest & " seconds" & vbNewLine
For i = 1 To n
    message = message & i & vbTab & solution(i + solLM1) & vbNewLine
Next i
MsgBox message

'ConjugateGradientFR
TimerStart = timer
For j = 1 To 1000
i = ConjugateGradientFR(xInitial(), dx(), solution(), 500, 10)
Next j
TimerEnd = timer
TimerFletcherReeves = TimerEnd - TimerStart
If i > 500 Then
    MsgBox "ConjugateGradientFR Not converged yet."
Else
    MsgBox "ConjugateGradientFR Converged after " & i & " steps."
End If
message = "Solution of ConjugateGradientLocalMin: " & vbTab & TimerFletcherReeves & " seconds" & vbNewLine
For i = 1 To n
    message = message & i & vbTab & solution(i + solLM1) & vbNewLine
Next i
MsgBox message

'ConjugateGradientPR
TimerStart = timer
For j = 1 To 1000
i = ConjugateGradientPR(xInitial(), dx(), solution(), 500, 10, 0.00001)
Next j
TimerEnd = timer
TimerPolakRibiere = TimerEnd - TimerStart
If i > 500 Then
    MsgBox "ConjugateGradientPR Not converged yet."
Else
    MsgBox "ConjugateGradientPR Converged after " & i & " steps."
End If
message = "Solution of ConjugateGradientPR: " & vbTab & TimerPolakRibiere & " seconds" & vbNewLine
For i = 1 To n
    message = message & i & vbTab & solution(i + solLM1) & vbNewLine
Next i
MsgBox message

'SimulatedAnnealing
'TempStart must be large enough to make the uphill and
'downhill transition probabilities about the same.
'exp(-abs(Fitness(xInitial+dx)-Fitness(xInitial))/TempStart) = 0.99
'MaxIteration should be more than 20*abs(xBest-xInitial)/dx
'larger MaxIteration won't improve the solution.
'TempStop should be as small as possible and prevent to
'cause "numerical overflow" at the same time.
dx(1 + dxLM1) = 0.1
dx(2 + dxLM1) = 0.1
TimerStart = timer
Call SimulatedAnnealing(xInitial(), dx(), _
                                 solution(), 2000, 40000#, 0.00001)
TimerEnd = timer
TimerSimulatedAnnealing = TimerEnd - TimerStart
message = "Solution of SimulatedAnnealing: " & vbTab & TimerSimulatedAnnealing & " seconds" & vbNewLine
For i = 1 To n
    message = message & i & vbTab & solution(i + solLM1) & vbNewLine
Next i
MsgBox message

'GeneticAlgorithm
Dim NumBit() As Integer
Dim UpperBound() As Double, LowerBound() As Double
Dim MateRate() As Double, MuteRate() As Double
Const BitLM1 As Integer = 53
Const UpperLM1 As Integer = 59
Const LowerLM1 As Integer = 61
Const MateLM1 As Integer = 67
Const MuteLM1 As Integer = 71
ReDim NumBit(1 + BitLM1 To n + BitLM1)
ReDim UpperBound(1 + UpperLM1 To n + UpperLM1)
ReDim LowerBound(1 + LowerLM1 To n + LowerLM1)
ReDim MateRate(1 + MateLM1 To n + MateLM1)
ReDim MuteRate(1 + MuteLM1 To n + MuteLM1)
For i = 1 To n
    NumBit(i + BitLM1) = 4&
    UpperBound(i + UpperLM1) = 1#
    LowerBound(i + LowerLM1) = -1#
    MateRate(i + MateLM1) = 0.6
    MuteRate(i + MuteLM1) = 0.01
Next i
TimerStart = timer
Call GeneticAlgorithm(100, 100, solution(), _
        NumBit(), UpperBound(), LowerBound(), _
        MateRate(), MuteRate())
TimerEnd = timer
TimerGeneticAlgorithm = TimerEnd - TimerStart
message = "Solution of GeneticAlgorithm: " & vbTab & TimerGeneticAlgorithm & " seconds" & vbNewLine
For i = 1 To n
    message = message & i & vbTab & solution(i + solLM1) & vbNewLine
Next i
MsgBox message
        
End Sub
'gradient(f) =0 at local maximum or minimum
'where f = f(x1, x2, x3, ..., xn)
'
'name of function f is "Fitness"
'Fitness has parameter: array x() As Double
'
'f1 = df/dx1 = 0
'f2 = df/dx2 = 0
'f3 = df/dx3 = 0
'...
'fn = df/dxn = 0
'
'f1 + df1/dx1*Dx1 + df1/dx2*Dx2 + ... + df1/dxn*Dxn = 0
'f2 + df2/dx1*Dx1 + df2/dx2*Dx2 + ... + df2/dxn*Dxn = 0
'f3 + df3/dx1*Dx1 + df3/dx2*Dx2 + ... + df3/dxn*Dxn = 0
'...
'fn + dfn/dx1*Dx1 + dfn/dx2*Dx2 + ... + dfn/dxn*Dxn = 0
'
'gradient(f) + Jacobian(f) * Dx = 0
'x = x + Dx
Public Function NewtonOptimization(xInitial() As Double, _
                       dx() As Double, _
                       solution() As Double, _
                       MaxIteration As Integer) As Integer
Dim i As Integer, j As Integer
Dim ii As Integer, jj As Integer
Dim IterationCounter As Integer
Dim iniLM1 As Integer, dxLM1 As Integer, solLM1 As Integer
Dim UpperBound As Integer
Dim tmp1 As Double, tmp2 As Double
Dim dxi As Double, dxj As Double
Dim xiPdxi As Double, xiMdxi As Double
Dim fc As Double
Dim fPP As Double, fMM As Double
Dim fPM As Double, fMP As Double
Dim dfNorm As Double
Dim tolerence As Double
Dim xtmp() As Double
Dim Mdf() As Double
Dim xStep() As Double
Dim dxInv() As Double
Dim dx2Inv() As Double
Dim Jacobian() As Double

iniLM1 = LBound(xInitial) - 1
dxLM1 = LBound(dx) - 1
solLM1 = LBound(solution) - 1
UpperBound = UBound(xInitial) - LBound(xInitial) + 1

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
ReDim Mdf(1 To UpperBound)
ReDim xStep(1 To UpperBound)
ReDim dxInv(1 To UpperBound)
ReDim dx2Inv(1 To UpperBound, 1 To UpperBound)
ReDim Jacobian(1 To UpperBound, 1 To UpperBound)
'initiate
For i = 1 To UpperBound
    ii = i + iniLM1
    xtmp(ii) = xInitial(ii)
Next i

tolerence = 0.000000001 * Sqr(CDbl(UpperBound))
dfNorm = 0#
fc = Fitness(xtmp())
For i = 1 To UpperBound
    dxi = dx(i + dxLM1)
'initiate dxInv(i) and dx2Inv(i, j)
    dxInv(i) = 1# / (2# * dxi)
    dx2Inv(i, i) = 1# / (dxi * dxi)
    For j = i + 1 To UpperBound
        tmp1 = 1# / (2# * dxi * 2# * dx(j + dxLM1))
        dx2Inv(i, j) = tmp1
        dx2Inv(j, i) = tmp1
    Next j
    ii = i + iniLM1
    tmp1 = xtmp(ii)
    xtmp(ii) = tmp1 + dxi
    fPP = Fitness(xtmp())
    xtmp(ii) = tmp1 - dxi
    fMM = Fitness(xtmp())
    xtmp(ii) = tmp1
'-df and Norm(central difference)
    tmp2 = (fPP - fMM) * dxInv(i)
    Mdf(i) = -tmp2
    dfNorm = dfNorm + tmp2 * tmp2
'diagonal element of Jacobian matrix
    Jacobian(i, i) = (fPP - 2# * fc + fMM) * dx2Inv(i, i)
Next i
dfNorm = Sqr(dfNorm)
IterationCounter = 0
Do While (dfNorm > tolerence)
    IterationCounter = IterationCounter + 1
    If IterationCounter > MaxIteration Then Exit Do
'off-diagonal element of Jacobian matrix
    For i = 1 To UpperBound - 1
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xiPdxi = tmp1 + dxi
        xiMdxi = tmp1 - dxi
        For j = i + 1 To UpperBound
            dxj = dx(j + dxLM1)
            jj = j + iniLM1
            tmp2 = xtmp(jj)

            xtmp(ii) = xiPdxi
            xtmp(jj) = tmp2 + dxj
            fPP = Fitness(xtmp())
            xtmp(ii) = xiMdxi
            fMP = Fitness(xtmp())
            xtmp(jj) = tmp2 - dxj
            fMM = Fitness(xtmp())
            xtmp(ii) = xiPdxi
            fPM = Fitness(xtmp())
            
            xtmp(jj) = tmp2
            tmp2 = (fPP - fMP - fPM + fMM) * dx2Inv(i, j)
            Jacobian(i, j) = tmp2
            Jacobian(j, i) = tmp2
        Next j
        xtmp(ii) = tmp1
    Next i
'call LU decomposition to solve
    Call LUSymmetricLinearSolver(Jacobian(), Mdf(), xStep())
'update x
    For i = 1 To UpperBound
        ii = i + iniLM1
        xtmp(ii) = xtmp(ii) + xStep(i)
    Next i
'Norm and diagonal element
    dfNorm = 0#
    fc = Fitness(xtmp())
    For i = 1 To UpperBound
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xtmp(ii) = tmp1 + dxi
        fPP = Fitness(xtmp())
        xtmp(ii) = tmp1 - dxi
        fMM = Fitness(xtmp())
        xtmp(ii) = tmp1
'-df and Norm(central difference)
        tmp2 = (fPP - fMM) * dxInv(i)
        Mdf(i) = -tmp2
        dfNorm = dfNorm + tmp2 * tmp2
'diagonal element of Jacobian matrix
        Jacobian(i, i) = (fPP - 2# * fc + fMM) * dx2Inv(i, i)
    Next i
    dfNorm = Sqr(dfNorm)
Loop

NewtonOptimization = IterationCounter
For i = 1 To UpperBound
    solution(i + solLM1) = xtmp(i + iniLM1)
Next i

End Function
'gradient(f) =0 at local maximum or minimum
'where f = f(x1, x2, x3, ..., xn)
'
'name of function f is "Fitness"
'Fitness has parameter: array x() As Double
'
'r = gradient(f): vector, residue
'
'x0: vector, previous solution
'r0 = (r1, r2, r3, ..., rn): vector, previous residue
'alpha: , step size
'new_x: vector, new solution
'
'since (-r0) is the steepest descent direction
'new_x = x0 + alpha * r0
'
'find alpha to maximize/minimize f
'df(new_X)/dalpha = 0
'gradient(f) dot r0 = 0
'(r0 dot r0) + alpha * ((Jacobian(f) * r0) dot r0) = 0
Public Function SteepestDescent(xInitial() As Double, _
                                dx() As Double, _
                                solution() As Double, _
                                MaxIteration As Integer) As Integer
Dim i As Integer, j As Integer
Dim ii As Integer, jj As Integer
Dim IterationCounter As Integer
Dim iniLM1 As Integer, dxLM1 As Integer, solLM1 As Integer
Dim UpperBound As Integer
Dim tmp1 As Double, tmp2 As Double
Dim dxi As Double, dxj As Double
Dim xiPdxi As Double, xiMdxi As Double
Dim fc As Double
Dim fPP As Double, fMM As Double
Dim fPM As Double, fMP As Double
Dim rSquare As Double
Dim tolerence As Double
Dim alpha As Double
Dim xtmp() As Double
Dim r() As Double
Dim rJacobian() As Double
Dim dxInv() As Double
Dim dx2Inv() As Double

iniLM1 = LBound(xInitial) - 1
dxLM1 = LBound(dx) - 1
solLM1 = LBound(solution) - 1
UpperBound = UBound(xInitial) - iniLM1

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
ReDim r(1 To UpperBound)
ReDim rJacobian(1 To UpperBound)
ReDim dxInv(1 To UpperBound)
ReDim dx2Inv(1 To UpperBound, 1 To UpperBound)
'initiate
For i = 1 To UpperBound
    ii = i + iniLM1
    xtmp(ii) = xInitial(ii)
Next i
tolerence = 0.000000001 * Sqr(CDbl(UpperBound))

rSquare = 0#
fc = Fitness(xtmp())
For i = 1 To UpperBound
    dxi = dx(i + dxLM1)
'initiate dxInv(i) and dx2Inv(i, j)
    dxInv(i) = 1# / (2# * dxi)
    dx2Inv(i, i) = 1# / (dxi * dxi)
    For j = i + 1 To UpperBound
        tmp1 = 1# / (2# * dxi * 2# * dx(j + dxLM1))
        dx2Inv(i, j) = tmp1
        dx2Inv(j, i) = tmp1
    Next j
    ii = i + iniLM1
    tmp1 = xtmp(ii)
    xtmp(ii) = tmp1 + dxi
    fPP = Fitness(xtmp())
    xtmp(ii) = tmp1 - dxi
    fMM = Fitness(xtmp())
    xtmp(ii) = tmp1
'r and rSquare(central difference)
    tmp2 = (fPP - fMM) * dxInv(i)
    r(i) = tmp2
    rSquare = rSquare + tmp2 * tmp2
'diagonal element of Jacobian matrix
    rJacobian(i) = (fPP - 2# * fc + fMM) * dx2Inv(i, i) _
                   * r(i)
Next i
IterationCounter = 0
Do While (Sqr(rSquare) > tolerence)
    IterationCounter = IterationCounter + 1
    If IterationCounter > MaxIteration Then Exit Do
'off-diagonal element of Jacobian matrix
    For i = 1 To UpperBound - 1
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xiPdxi = tmp1 + dxi
        xiMdxi = tmp1 - dxi
        For j = i + 1 To UpperBound
            dxj = dx(j + dxLM1)
            jj = j + iniLM1
            tmp2 = xtmp(jj)

            xtmp(ii) = xiPdxi
            xtmp(jj) = tmp2 + dxj
            fPP = Fitness(xtmp())
            xtmp(ii) = xiMdxi
            fMP = Fitness(xtmp())
            xtmp(jj) = tmp2 - dxj
            fMM = Fitness(xtmp())
            xtmp(ii) = xiPdxi
            fPM = Fitness(xtmp())
            
            xtmp(jj) = tmp2
            tmp2 = (fPP - fMP - fPM + fMM) * dx2Inv(i, j)
            rJacobian(i) = rJacobian(i) + tmp2 * r(j)
            rJacobian(j) = rJacobian(j) + tmp2 * r(i)
        Next j
        xtmp(ii) = tmp1
    Next i
'(Jacobian(f) * r0) dot r0) and alpha
    tmp1 = 0#
    For i = 1 To UpperBound
        tmp1 = tmp1 + rJacobian(i) * r(i)
    Next i
    alpha = -rSquare / tmp1
'update x
    For i = 1 To UpperBound
        ii = i + iniLM1
        xtmp(ii) = xtmp(ii) + alpha * r(i)
    Next i
'diagonal element of Jacobian matrix
    rSquare = 0#
    fc = Fitness(xtmp())
    For i = 1 To UpperBound
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xtmp(ii) = tmp1 + dxi
        fPP = Fitness(xtmp())
        xtmp(ii) = tmp1 - dxi
        fMM = Fitness(xtmp())
        xtmp(ii) = tmp1
'r and rSquare(central difference)
        tmp2 = (fPP - fMM) * dxInv(i)
        r(i) = tmp2
        rSquare = rSquare + tmp2 * tmp2
'diagonal element of Jacobian matrix
        rJacobian(i) = (fPP - 2# * fc + fMM) * dx2Inv(i, i) _
                          * r(i)
    Next i
Loop

SteepestDescent = IterationCounter
For i = 1 To UpperBound
    solution(i + solLM1) = xtmp(i + iniLM1)
Next i

End Function
'Nonlinear conjugate gradients with
'Newton-Raphson and Fletcher-Reeves
'
'name of function f is "Fitness"
'Fitness has parameter: array x() As Double
Public Function ConjugateGradientFR(xInitial() As Double, _
                                    dx() As Double, _
                                    solution() As Double, _
                                    MaxIteration As Integer, _
                                    MaxInnerIteration As Integer) As Integer
Dim i As Integer, j As Integer
Dim k As Integer
Dim ii As Integer, jj As Integer
Dim IterationCounter As Integer
Dim InnerIterationCounter As Integer
Dim iniLM1 As Integer, dxLM1 As Integer, solLM1 As Integer
Dim UpperBound As Integer
Dim tmp1 As Double, tmp2 As Double
Dim dxi As Double, dxj As Double
Dim xiPdxi As Double, xiMdxi As Double
Dim fc As Double
Dim fPP As Double, fMM As Double
Dim fPM As Double, fMP As Double
Dim rSquareNew As Double, rSquareOld As Double
Dim dSquare As Double
Dim rTolerence As Double
Dim dTolerence As Double
Dim numerator As Double
Dim alpha As Double, beta As Double
Dim xtmp() As Double
Dim r() As Double
Dim d() As Double
Dim dJacobian() As Double
Dim dxInv() As Double
Dim dx2Inv() As Double

iniLM1 = LBound(xInitial) - 1
dxLM1 = LBound(dx) - 1
solLM1 = LBound(solution) - 1
UpperBound = UBound(xInitial) - iniLM1

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
ReDim r(1 To UpperBound)
ReDim d(1 To UpperBound)
ReDim dJacobian(1 To UpperBound)
ReDim dxInv(1 To UpperBound)
ReDim dx2Inv(1 To UpperBound, 1 To UpperBound)
'initiate
For i = 1 To UpperBound
    ii = i + iniLM1
    xtmp(ii) = xInitial(ii)
Next i
rTolerence = 0.000000001 * Sqr(CDbl(UpperBound))
dTolerence = 0.000000001 * Sqr(CDbl(UpperBound))

rSquareNew = 0#
fc = Fitness(xtmp())
For i = 1 To UpperBound
    dxi = dx(i + dxLM1)
'initiate dxInv(i) and dx2Inv(i, j)
    dxInv(i) = 1# / (2# * dxi)
    dx2Inv(i, i) = 1# / (dxi * dxi)
    For j = i + 1 To UpperBound
        tmp1 = 1# / (2# * dxi * 2# * dx(j + dxLM1))
        dx2Inv(i, j) = tmp1
        dx2Inv(j, i) = tmp1
    Next j
    ii = i + iniLM1
    tmp1 = xtmp(ii)
    xtmp(ii) = tmp1 + dxi
    fPP = Fitness(xtmp())
    xtmp(ii) = tmp1 - dxi
    fMM = Fitness(xtmp())
    xtmp(ii) = tmp1
'r, d and rSquare(central difference)
    tmp2 = -(fPP - fMM) * dxInv(i)
    r(i) = tmp2
    d(i) = tmp2
    rSquareNew = rSquareNew + tmp2 * tmp2
Next i
IterationCounter = 0
k = 0
Do While (Sqr(rSquareNew) > rTolerence)
    IterationCounter = IterationCounter + 1
    If IterationCounter > MaxIteration Then Exit Do
'Newton-Raphson inner loop
    InnerIterationCounter = 0
    dSquare = 0#
    For i = 1 To UpperBound
        tmp1 = d(i)
        dSquare = dSquare + tmp1 * tmp1
    Next i
    Do
        InnerIterationCounter = InnerIterationCounter + 1
        If InnerIterationCounter > MaxInnerIteration Then Exit Do
'Norm and diagonal element of Jacobian matrix
        numerator = 0#
        fc = Fitness(xtmp())
        For i = 1 To UpperBound
            dxi = dx(i + dxLM1)
            ii = i + iniLM1
            tmp1 = xtmp(ii)
            xtmp(ii) = tmp1 + dxi
            fPP = Fitness(xtmp())
            xtmp(ii) = tmp1 - dxi
            fMM = Fitness(xtmp())
            xtmp(ii) = tmp1
'(gradient(f) dot d)(central difference)
            numerator = numerator + (fPP - fMM) * dxInv(i) * d(i)
'diagonal element of Jacobian matrix
            dJacobian(i) = (fPP - 2# * fc + fMM) * dx2Inv(i, i) _
                              * d(i)
        Next i
'off-diagonal element of Jacobian matrix
        For i = 1 To UpperBound - 1
            dxi = dx(i + dxLM1)
            ii = i + iniLM1
            tmp1 = xtmp(ii)
            xiPdxi = tmp1 + dxi
            xiMdxi = tmp1 - dxi
            For j = i + 1 To UpperBound
                dxj = dx(j + dxLM1)
                jj = j + iniLM1
                tmp2 = xtmp(jj)
    
                xtmp(ii) = xiPdxi
                xtmp(jj) = tmp2 + dxj
                fPP = Fitness(xtmp())
                xtmp(ii) = xiMdxi
                fMP = Fitness(xtmp())
                xtmp(jj) = tmp2 - dxj
                fMM = Fitness(xtmp())
                xtmp(ii) = xiPdxi
                fPM = Fitness(xtmp())
                
                xtmp(jj) = tmp2
                tmp2 = (fPP - fMP - fPM + fMM) * dx2Inv(i, j)
                dJacobian(i) = dJacobian(i) + tmp2 * d(j)
                dJacobian(j) = dJacobian(j) + tmp2 * d(i)
            Next j
            xtmp(ii) = tmp1
        Next i
'(Jacobian(f) * r0) dot r0) and alpha
        tmp1 = 0#
        For i = 1 To UpperBound
            tmp1 = tmp1 + dJacobian(i) * d(i)
        Next i
        alpha = -numerator / tmp1
        For i = 1 To UpperBound
            ii = i + iniLM1
            xtmp(ii) = xtmp(ii) + alpha * d(i)
        Next i
    Loop While (alpha * Sqr(dSquare) > dTolerence)
'r and rSquare(central difference)
    rSquareOld = rSquareNew
    rSquareNew = 0#
    fc = Fitness(xtmp())
    For i = 1 To UpperBound
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xtmp(ii) = tmp1 + dxi
        fPP = Fitness(xtmp())
        xtmp(ii) = tmp1 - dxi
        fMM = Fitness(xtmp())
        xtmp(ii) = tmp1
'r and rSquare(central difference)
        tmp2 = -(fPP - fMM) * dxInv(i)
        r(i) = tmp2
        rSquareNew = rSquareNew + tmp2 * tmp2
    Next i
'update d
    beta = rSquareNew / rSquareOld
    tmp1 = 0#
    For i = 1 To UpperBound
        d(i) = r(i) + beta * d(i)
        tmp1 = tmp1 + r(i) * d(i)
    Next i
'restart
    k = k + 1
    If k = 50 Or tmp1 <= 0# Then
        k = 0
        For i = 1 To UpperBound
            d(i) = r(i)
        Next i
    End If
Loop

ConjugateGradientFR = IterationCounter
For i = 1 To UpperBound
    solution(i + solLM1) = xtmp(i + iniLM1)
Next i

End Function
'Preconditioned nonlinear conjugate gradients
'with secant and Polak-Ribiere
'
'name of function f is "Fitness"
'Fitness has parameter: array x() As Double
Public Function ConjugateGradientPR(xInitial() As Double, _
                                    dx() As Double, _
                                    solution() As Double, _
                                    MaxIteration As Integer, _
                                    MaxInnerIteration As Integer, _
                                    alphaInit As Double) As Integer
Dim i As Integer, j As Integer
Dim k As Integer
Dim ii As Integer, jj As Integer
Dim IterationCounter As Integer
Dim InnerIterationCounter As Integer
Dim iniLM1 As Integer, dxLM1 As Integer, solLM1 As Integer
Dim UpperBound As Integer
Dim tmp1 As Double, tmp2 As Double
Dim dxi As Double, dxj As Double
Dim xiPdxi As Double, xiMdxi As Double
Dim fc As Double
Dim fPP As Double, fMM As Double
Dim fPM As Double, fMP As Double
Dim deltaNew As Double, deltaMid As Double, deltaOld As Double
Dim dSquare As Double
Dim deltaTolerence As Double
Dim dTolerence As Double
Dim alpha As Double, beta As Double
Dim gamma As Double, gammaPrev As Double
Dim xtmp() As Double
Dim r() As Double
Dim s() As Double
Dim d() As Double
Dim dxInv() As Double
Dim dx2Inv() As Double
Dim m() As Double

iniLM1 = LBound(xInitial) - 1
dxLM1 = LBound(dx) - 1
solLM1 = LBound(solution) - 1
UpperBound = UBound(xInitial) - iniLM1

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
ReDim r(1 To UpperBound)
ReDim s(1 To UpperBound)
ReDim d(1 To UpperBound)
ReDim dxInv(1 To UpperBound)
ReDim dx2Inv(1 To UpperBound, 1 To UpperBound)
ReDim m(1 To UpperBound, 1 To UpperBound)
'initiate
For i = 1 To UpperBound
    ii = i + iniLM1
    xtmp(ii) = xInitial(ii)
Next i
deltaTolerence = 0.000000001 * Sqr(CDbl(UpperBound))
dTolerence = 0.000000001 * Sqr(CDbl(UpperBound))

'preconditioner M = Jacobian
fc = Fitness(xtmp())
For i = 1 To UpperBound
    dxi = dx(i + dxLM1)
'initiate dxInv(i) and dx2Inv(i, j)
    dxInv(i) = 1# / (2# * dxi)
    dx2Inv(i, i) = 1# / (dxi * dxi)
    For j = i + 1 To UpperBound
        tmp1 = 1# / (2# * dxi * 2# * dx(j + dxLM1))
        dx2Inv(i, j) = tmp1
        dx2Inv(j, i) = tmp1
    Next j
    ii = i + iniLM1
    tmp1 = xtmp(ii)
    xtmp(ii) = tmp1 + dxi
    fPP = Fitness(xtmp())
    xtmp(ii) = tmp1 - dxi
    fMM = Fitness(xtmp())
    xtmp(ii) = tmp1
'r(central difference)
    r(i) = -(fPP - fMM) * dxInv(i)
'diagonal element of Jacobian matrix
    m(i, i) = (fPP - 2# * fc + fMM) * dx2Inv(i, i)
Next i
'off-diagonal element of Jacobian matrix
For i = 1 To UpperBound - 1
    dxi = dx(i + dxLM1)
    ii = i + iniLM1
    tmp1 = xtmp(ii)
    xiPdxi = tmp1 + dxi
    xiMdxi = tmp1 - dxi
    For j = i + 1 To UpperBound
        dxj = dx(j + dxLM1)
        jj = j + iniLM1
        tmp2 = xtmp(jj)

        xtmp(ii) = xiPdxi
        xtmp(jj) = tmp2 + dxj
        fPP = Fitness(xtmp())
        xtmp(ii) = xiMdxi
        fMP = Fitness(xtmp())
        xtmp(jj) = tmp2 - dxj
        fMM = Fitness(xtmp())
        xtmp(ii) = xiPdxi
        fPM = Fitness(xtmp())
        
        xtmp(jj) = tmp2
        tmp2 = (fPP - fMP - fPM + fMM) * dx2Inv(i, j)
        m(i, j) = tmp2
        m(j, i) = tmp2
    Next j
    xtmp(ii) = tmp1
Next i
Call LUSymmetricLinearSolver(m(), r(), s())
deltaNew = 0#
For i = 1 To UpperBound
    d(i) = s(i)
    deltaNew = deltaNew + r(i) * d(i)
Next i
IterationCounter = 0
k = 0
Do While (Sqr(deltaNew) > deltaTolerence)
    IterationCounter = IterationCounter + 1
    If IterationCounter > MaxIteration Then Exit Do
'Secant inner loop
    InnerIterationCounter = 0
    dSquare = 0#
    For i = 1 To UpperBound
        tmp1 = d(i)
        dSquare = dSquare + tmp1 * tmp1
    Next i
    alpha = -alphaInit
    For i = 1 To UpperBound
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        solution(i + solLM1) = tmp1
        xtmp(ii) = tmp1 + alphaInit * d(i)
    Next i
'gammaPrev
    gammaPrev = 0#
    fc = Fitness(xtmp())
    For i = 1 To UpperBound
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xtmp(ii) = tmp1 + dxi
        fPP = Fitness(xtmp())
        xtmp(ii) = tmp1 - dxi
        fMM = Fitness(xtmp())
        xtmp(ii) = tmp1
'(gradient(f) dot d)(central difference)
        gammaPrev = gammaPrev + (fPP - fMM) * dxInv(i) * d(i)
    Next i
    For i = 1 To UpperBound
        xtmp(i + iniLM1) = solution(i + solLM1)
    Next i
    Do
        InnerIterationCounter = InnerIterationCounter + 1
        If InnerIterationCounter > MaxInnerIteration Then Exit Do
'gamma
        gamma = 0#
        fc = Fitness(xtmp())
        For i = 1 To UpperBound
            dxi = dx(i + dxLM1)
            ii = i + iniLM1
            tmp1 = xtmp(ii)
            xtmp(ii) = tmp1 + dxi
            fPP = Fitness(xtmp())
            xtmp(ii) = tmp1 - dxi
            fMM = Fitness(xtmp())
            xtmp(ii) = tmp1
'(gradient(f) dot d)(central difference)
            gamma = gamma + (fPP - fMM) * dxInv(i) * d(i)
        Next i
        alpha = alpha * gamma / (gammaPrev - gamma)
        For i = 1 To UpperBound
            ii = i + iniLM1
            xtmp(ii) = xtmp(ii) + alpha * d(i)
        Next i
        gammaPrev = gamma
    Loop While (alpha * Sqr(dSquare) > dTolerence)
'deltaOld,r and deltaMid(central difference)
    deltaOld = deltaNew
    deltaMid = 0#
    fc = Fitness(xtmp())
    For i = 1 To UpperBound
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xtmp(ii) = tmp1 + dxi
        fPP = Fitness(xtmp())
        xtmp(ii) = tmp1 - dxi
        fMM = Fitness(xtmp())
        xtmp(ii) = tmp1
'r and deltaMid(central difference)
        tmp2 = -(fPP - fMM) * dxInv(i)
        r(i) = tmp2
        deltaMid = deltaMid + tmp2 * s(i)
'diagonal element of Jacobian matrix
        m(i, i) = (fPP - 2# * fc + fMM) * dx2Inv(i, i)
    Next i
'off-diagonal element of Jacobian matrix
    For i = 1 To UpperBound - 1
        dxi = dx(i + dxLM1)
        ii = i + iniLM1
        tmp1 = xtmp(ii)
        xiPdxi = tmp1 + dxi
        xiMdxi = tmp1 - dxi
        For j = i + 1 To UpperBound
            dxj = dx(j + dxLM1)
            jj = j + iniLM1
            tmp2 = xtmp(jj)
    
            xtmp(ii) = xiPdxi
            xtmp(jj) = tmp2 + dxj
            fPP = Fitness(xtmp())
            xtmp(ii) = xiMdxi
            fMP = Fitness(xtmp())
            xtmp(jj) = tmp2 - dxj
            fMM = Fitness(xtmp())
            xtmp(ii) = xiPdxi
            fPM = Fitness(xtmp())
            
            xtmp(jj) = tmp2
            tmp2 = (fPP - fMP - fPM + fMM) * dx2Inv(i, j)
            m(i, j) = tmp2
            m(j, i) = tmp2
        Next j
        xtmp(ii) = tmp1
    Next i
    Call LUSymmetricLinearSolver(m(), r(), s())
    deltaNew = 0#
    For i = 1 To UpperBound
        deltaNew = deltaNew + r(i) * s(i)
    Next i
    beta = (deltaNew - deltaMid) / deltaOld
    k = k + 1
'restart
    If k = 50 Or beta <= 0# Then
        k = 0
        For i = 1 To UpperBound
            d(i) = s(i)
        Next i
'update d
    Else
        For i = 1 To UpperBound
            d(i) = s(i) + beta * d(i)
        Next i
    End If
Loop

ConjugateGradientPR = IterationCounter
For i = 1 To UpperBound
    solution(i + solLM1) = xtmp(i + iniLM1)
Next i

End Function
'£_E = FunctionNew - FunctionOld
'if £_E <= 0, downhill moves are always performed;
'if £_E > 0, the probability of performing uphill moves is
'the Maxwell-Boltzmann distribution(exp(-£_E/T)).
'where T is the absolute temperature(T > 0) and T decreases.
'For global minimum only.
'
'name of target function is "Fitness"
'Fitness has parameter: array x() As Double
Public Sub SimulatedAnnealing(xInitial() As Double, _
                              dx() As Double, _
                              solution() As Double, _
                              MaxIteration As Integer, _
                              TempStart As Double, _
                              TempStop As Double)
Dim IterationCounter As Integer
Dim dxLM1 As Integer, solLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer, ii As Integer
Dim TempFactor As Double
Dim Temp As Double, dTemp As Double
Dim FunctionOld As Double, FunctionNew As Double
Dim xtmp() As Double

dxLM1 = LBound(dx) - 1
solLM1 = LBound(solution) - 1
UpperBound = UBound(solution) - solLM1

If UpperBound <> (UBound(dx) - dxLM1) Then
    MsgBox "Number of element of Parameter array solution() " & vbNewLine & _
           "is not equal to that of array dx()!"
    Exit Sub
End If
If UpperBound <> (UBound(xInitial) - LBound(xInitial) + 1) Then
    MsgBox "Number of element of Parameter array solution() " & vbNewLine & _
           "is not equal to that of array xInitial()!"
    Exit Sub
End If
If TempStart <= 0# Or TempStop <= 0# Then
    MsgBox "Parameter TempStart/TempStop should be greater than zero!"
    Exit Sub
End If
If TempStart <= TempStop Then
    MsgBox "Parameter TempStart should be greater than parameter TempStop!"
    Exit Sub
End If

ReDim xtmp(1 + solLM1 To UpperBound + solLM1)
ii = LBound(xInitial)
For i = 1 To UpperBound
    solution(i + solLM1) = xInitial(i + ii - 1)
Next i
Temp = TempStart
TempFactor = Exp(Log(TempStop / TempStart) / CDbl(MaxIteration - 1))
Randomize timer

FunctionOld = Fitness(solution())
For IterationCounter = 1 To MaxIteration
    For i = 1 To UpperBound
        ii = i + solLM1
        xtmp(ii) = solution(ii) + 2# * (Rnd - 0.5) * dx(i + dxLM1)
    Next i
    FunctionNew = Fitness(xtmp())
    dTemp = FunctionNew - FunctionOld
    If (dTemp <= 0#) Then
        For i = 1 To UpperBound
            ii = i + solLM1
            solution(ii) = xtmp(ii)
        Next i
    ElseIf (Rnd < Exp(-dTemp / Temp)) Then
        For i = 1 To UpperBound
            ii = i + solLM1
            solution(ii) = xtmp(ii)
        Next i
    End If
'Dim sh As Worksheet
'Set sh = ThisWorkbook.Worksheets("SimulatedAnnealing")
'For i = 1 To UpperBound
'sh.Cells(IterationCounter, i).Value = solution(i + solLM1)
'Next i
'sh.Cells(IterationCounter, i).Value = Temp
    FunctionOld = FunctionNew
    Temp = Temp * TempFactor
Next IterationCounter

End Sub
'initialize chromosome randomly
'For each generation, process the following three steps:
'Selection Step
'Crossover Step
'Mutation Step
'For global maximum only.
'
'name of target function is "Fitness"
'Fitness has parameter: array x() As Double
Public Sub GeneticAlgorithm(NumGeneration As Long, _
        NumPopulation As Long, solution() As Double, _
        NumBit() As Integer, UpperBound() As Double, _
        LowerBound() As Double, MateRate() As Double, _
        MuteRate() As Double)
Dim i As Integer, j As Integer
Dim k As Integer, l As Integer
Dim NumVariable As Integer
Dim solLM1 As Integer, PopLM1 As Integer
Dim BitLM1 As Integer, UppLM1 As Integer
Dim LowLM1 As Integer, MatLM1 As Integer
Dim MutLM1 As Integer
Dim CrossOver1 As Integer, CrossOver2 As Integer
Dim Index1 As Integer, Index2 As Integer
Dim chromosomeTmp As String
Dim chromosome() As String
Dim GenerationCounter As Long
Dim li As Long, lj As Long
Dim lk As Long, ll As Long
Dim TwoNM1() As Long
Dim IndexArray() As Long
Dim BitValue(1 To 31) As Long
Dim NumMate() As Long
Dim tmp As Double
Dim fMax As Double, fMin As Double
Dim fIntervalInv As Double
Dim ProbilityRatio As Double
Const LowerProbility As Double = 0.4
Const UpperProbility As Double = 0.6
Dim f() As Double
Dim IntervalRatio() As Double
Dim LoopOver As Boolean
Dim mated() As Boolean

solLM1 = LBound(solution) - 1
BitLM1 = LBound(NumBit) - 1
UppLM1 = LBound(UpperBound) - 1
LowLM1 = LBound(LowerBound) - 1
MatLM1 = LBound(MateRate) - 1
MutLM1 = LBound(MuteRate) - 1
NumVariable = UBound(solution) - solLM1

If LowerProbility <= 0# Or LowerProbility >= 1# Then
    MsgBox "LowerProbility should be between 0.0 and 1.0!"
    Exit Sub
End If
If UpperProbility <= 0# Or UpperProbility >= 1# Then
    MsgBox "UpperProbility should be between 0.0 and 1.0!"
    Exit Sub
End If
If UpperProbility < LowerProbility Then
    MsgBox "UpperProbility should be greater than LowerProbility!"
    Exit Sub
End If
If NumVariable <> (UBound(NumBit) - BitLM1) Then
    MsgBox "Size of parameter array ""solution"" is not" & _
           "equal to that of parameter array ""NumBit""!"
    Exit Sub
End If
If NumVariable <> (UBound(UpperBound) - UppLM1) Then
    MsgBox "Size of parameter array ""solution"" is not" & _
           "equal to that of parameter array ""UpperBound""!"
    Exit Sub
End If
If NumVariable <> (UBound(LowerBound) - LowLM1) Then
    MsgBox "Size of parameter array ""solution"" is not" & _
           "equal to that of parameter array ""LowerBound""!"
    Exit Sub
End If
If NumVariable <> (UBound(MateRate) - MatLM1) Then
    MsgBox "Size of parameter array ""solution"" is not" & _
           "equal to that of parameter array ""MateRate""!"
    Exit Sub
End If
If NumVariable <> (UBound(MuteRate) - MutLM1) Then
    MsgBox "Size of parameter array ""solution"" is not" & _
           "equal to that of parameter array ""MuteRate""!"
    Exit Sub
End If

If NumGeneration < 1 Then
    MsgBox "Value of parameter ""NumGeneration"" should be greater than zero!"
    Exit Sub
End If
tmp = 0#
For i = LBound(NumBit) To UBound(NumBit)
    k = NumBit(i)
    If k > 31 Or k < 1 Then
        MsgBox "Value of Parameter array ""NumBit"" should be between 1 and 31!"
        Exit Sub
    End If
    tmp = tmp + CDbl(k) * Log(2#)
Next i
If NumPopulation < 2& Then
    MsgBox "Value of Parameter ""NumPopulation"" should be greater than or equal to 2!"
    Exit Sub
End If
If Log(CDbl(NumPopulation)) > tmp Then
    MsgBox "Value of Parameter array ""NumPopulation"" " & _
           "should be less than or equal to " & Exp(tmp)
    Exit Sub
End If
For i = LBound(MateRate) To UBound(MateRate)
    If MateRate(i) >= 1# Or MateRate(i) < 0# Then
        MsgBox "Value of Parameter array ""MateRate"" should " & _
               "be greater than or equal to 0 and less than 1!"
        Exit Sub
    End If
Next i
For i = LBound(MuteRate) To UBound(MuteRate)
    If MuteRate(i) > 0.01 Or MuteRate(i) < 0# Then
        MsgBox "Value of Parameter array ""MuteRate"" should be between 0 and 0.01!"
        Exit Sub
    End If
Next i

ReDim chromosome(1 To NumPopulation, 1 To NumVariable)
ReDim IndexArray(1 To NumPopulation)
ReDim NumMate(1 To NumVariable)
ReDim mated(1 To NumPopulation)
ReDim f(1 To NumPopulation)
ReDim IntervalRatio(1 To NumVariable)

BitValue(1) = 1&
For i = 2 To 31
    BitValue(i) = BitValue(i - 1) * 2&
Next i
''initialize chromosome randomly
Randomize timer
For li = 1 To NumPopulation
    For j = 1 To NumVariable
        chromosome(li, j) = ""
        For k = 1 To NumBit(j + BitLM1)
            If Rnd > 0.5 Then
                chromosome(li, j) = chromosome(li, j) & "1"
            Else
                chromosome(li, j) = chromosome(li, j) & "0"
            End If
        Next k
    Next j
Next li
For j = 1 To NumVariable
    NumMate(j) = CLng(MateRate(j + MatLM1) * CDbl(NumPopulation))
    If NumMate(j) Mod 2& <> 0& Then NumMate(j) = NumMate(j) - 1&
    IntervalRatio(j) = (UpperBound(j + UppLM1) - LowerBound(j + LowLM1)) / _
                       CDbl(CLng(2& ^ NumBit(j + BitLM1)) - 1&)
    If Len(chromosome(1, j)) <> NumBit(j + BitLM1) Then
        MsgBox "Number of bit is not the same as user-specified value" & _
               "for the " & j & "th variable."
    End If
Next j

ProbilityRatio = (UpperProbility - LowerProbility) / CDbl(NumGeneration - 1&)
For GenerationCounter = 1& To NumGeneration
''Selection Step
    For li = 1& To NumPopulation
        For j = 1 To NumVariable
            lj = 0&
            For k = 1 To NumBit(j + BitLM1)
                chromosomeTmp = mid(chromosome(li, j), k, 1)
                If chromosomeTmp = "1" Then
                    lj = lj + BitValue(k)
                ElseIf chromosomeTmp = "0" Then
                
                Else
                    MsgBox "Gene is out of range!" & vbNewLine & _
                           "at " & j & "th variable of " & li & "th chromosome."
                    Exit Sub
                End If
            Next k
            solution(j + solLM1) = CDbl(lj) * IntervalRatio(j) + LowerBound(j + LowLM1)
        Next j
        f(li) = Fitness(solution())
    Next li
    Call QuickPermutation(f(), IndexArray())
    tmp = LowerProbility + CDbl(GenerationCounter - 1&) * ProbilityRatio
    ll = CLng(tmp * NumPopulation)
    For li = 1 To ll
        lj = IndexArray(Int((NumPopulation - ll) * Rnd + ll + 1&))
        lk = IndexArray(li)
        For j = 1 To NumVariable
            chromosome(lk, j) = chromosome(lj, j)
        Next j
    Next li
''Crossover Step
    For i = 1 To NumVariable
        For lj = 1& To NumPopulation
            mated(lj) = False
        Next lj
        For lj = 1& To NumMate(i)
            Do
                ll = 1& + Int(Rnd * CDbl(NumPopulation))
                If mated(ll) = False Then
                    mated(ll) = True
                    IndexArray(lj) = ll
                    Exit Do
                End If
            Loop
        Next lj
        k = NumBit(i + BitLM1)
        For lj = 1& To (NumMate(i) - 1&) Step 2&
''two point crossover
            CrossOver1 = 1 + Int(Rnd * k)
            CrossOver2 = 1 + Int(Rnd * k)
''CrossOver2 is larger than CrossOver1
            If CrossOver1 > CrossOver2 Then
                l = CrossOver1
                CrossOver1 = CrossOver2
                CrossOver2 = l
            End If
            Index1 = IndexArray(lj)
            Index2 = IndexArray(lj + 1&)
            l = CrossOver2 - CrossOver1 + 1
            chromosomeTmp = mid(chromosome(Index1, i), CrossOver1, l)
            Mid(chromosome(Index1, i), CrossOver1, l) = _
                    mid(chromosome(Index2, i), CrossOver1, l)
            Mid(chromosome(Index2, i), CrossOver1, l) = _
                    chromosomeTmp
        Next lj
    Next i
''Mutation Step
    For i = 1 To NumVariable
        tmp = MuteRate(i + MutLM1)
        For lj = 1& To NumPopulation
            For k = 1 To NumBit(i + BitLM1)
                If Rnd < tmp Then
                    chromosomeTmp = mid(chromosome(lj, i), k, 1)
                    If chromosomeTmp = "1" Then
                        Mid(chromosome(lj, i), k, 1) = "0"
                    ElseIf chromosomeTmp = "0" Then
                        Mid(chromosome(lj, i), k, 1) = "1"
                    Else
                        MsgBox "Gene is out of range!" & vbNewLine & _
                               "at " & i & "th variable of " & j & "th chromosome."
                        Exit Sub
                    End If
                End If
            Next k
        Next lj
    Next i
Next GenerationCounter

For li = 1& To NumPopulation
    For j = 1 To NumVariable
        lj = 0&
        For k = 1 To NumBit(j + BitLM1)
            chromosomeTmp = mid(chromosome(li, j), k, 1)
            If chromosomeTmp = "1" Then
                lj = lj + BitValue(k)
            ElseIf chromosomeTmp = "0" Then
            
            Else
                MsgBox "Gene is out of range!" & vbNewLine & _
                       "at " & j & "th variable of " & li & "th chromosome."
            End If
        Next k
        solution(j + solLM1) = CDbl(lj) * IntervalRatio(j) + LowerBound(j + LowLM1)
    Next j
    f(li) = Fitness(solution())
Next li
lk = 1&
fMax = f(1&)
For li = 2& To NumPopulation
    tmp = f(li)
    If tmp > fMax Then
        fMax = tmp
        lk = li
    End If
Next li
For j = 1 To NumVariable
    lj = 0&
    For k = 1 To NumBit(j + BitLM1)
        chromosomeTmp = mid(chromosome(lk, j), k, 1)
        If chromosomeTmp = "1" Then
            lj = lj + BitValue(k)
        ElseIf chromosomeTmp = "0" Then
        
        Else
            MsgBox "Gene is out of range!" & vbNewLine & _
            "at " & j & "th variable of " & li & "th chromosome."
        End If
    Next k
    solution(j + solLM1) = CDbl(lj) * IntervalRatio(j) + LowerBound(j + LowLM1)
Next j
End Sub
'Target function for optimization
'Fitness has parameter: array x() As Double
Public Function Fitness(x() As Double) As Double
Dim x01 As Double, x02 As Double
Dim LowerM1 As Integer

LowerM1 = LBound(x) - 1
x01 = x(LowerM1 + 1)
x02 = x(LowerM1 + 2)

Fitness = 1# + x01 * x01 + x02 * x02
End Function
