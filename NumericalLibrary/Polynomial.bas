Attribute VB_Name = "Polynomial"
'Since "redim array" consumes lots of time, hence
'"redim array" is prevented as long as possible.
'And finding the true upper bound and true lower bound
'of an array is necessary.
Option Explicit
'
'
Function ShowPoly(poly() As Double) As String
Dim i As Integer

ShowPoly = ""
For i = LBound(poly) To UBound(poly)
    ShowPoly = ShowPoly & i & vbTab & poly(i) & vbNewLine
Next i

End Function
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub TestPolynomial()
Dim Poly1() As Double, Poly2() As Double
Dim polyA() As Double, polyB() As Double
Dim x As Double
Dim message As String
Dim i As Integer

ReDim Poly1(0 To 5)
ReDim Poly2(0 To 2)
'polyA and polyB is the answer array for polynomial operation
'redim answer array to arbitrary size
'polynomial operation will redim them to appropriate size
ReDim polyA(0 To 5)
ReDim polyB(0 To 5)
'(x - 1) * (x - 2) * (x - 3)
'= 1*x^3 -6*x^2 +11*x -6
Poly1(0) = -6
Poly1(1) = 11
Poly1(2) = -6
Poly1(3) = 1
Poly1(4) = 0
Poly1(5) = 0
'(x - 1) * (x - 2) * (x - 3) * (x - 4)
'= 1*x^4 -10*x^3 +35*x^2 -50*x +24
Poly2(0) = 0
Poly2(1) = 1
Poly2(2) = 2
'Poly2(3) = -10
'Poly2(4) = 1

'polyA() = (d/dx)poly1()
Call PolynomialDifferentiate(Poly1(), polyA())
message = "polyA() = (d/dx)poly1()" & vbNewLine
message = message & "poly1:" & vbNewLine
message = message & ShowPoly(Poly1())
message = message & vbNewLine
message = message & "polyA" & vbNewLine
message = message & ShowPoly(polyA())
message = message & vbNewLine
MsgBox message

'polyA() = ¡ìpoly1()dx
Call PolynomialIntegrate(Poly1(), polyA(), 1#)
message = "polyA() = ¡ìpoly1()dx" & vbNewLine
message = message & "poly1:" & vbNewLine
message = message & ShowPoly(Poly1())
message = message & vbNewLine
message = message & "polyA" & vbNewLine
message = message & ShowPoly(polyA())
message = message & vbNewLine
MsgBox message

'polyA() = poly1() + poly2()
Call PolynomialAdd(Poly1(), Poly2(), polyA())
message = "polyA() = poly1() + poly2()" & vbNewLine
message = message & "poly1:" & vbNewLine
message = message & ShowPoly(Poly1())
message = message & vbNewLine
message = message & "poly2:" & vbNewLine
message = message & ShowPoly(Poly2())
message = message & vbNewLine
message = message & "polyA" & vbNewLine
message = message & ShowPoly(polyA())
message = message & vbNewLine
MsgBox message

'polyA() = poly1() - poly2()
Call PolynomialSubtract(Poly1(), Poly2(), polyA())
message = "polyA() = poly1() - poly2()" & vbNewLine
message = message & "poly1:" & vbNewLine
message = message & ShowPoly(Poly1())
message = message & vbNewLine
message = message & "poly2:" & vbNewLine
message = message & ShowPoly(Poly2())
message = message & vbNewLine
message = message & "polyA" & vbNewLine
message = message & ShowPoly(polyA())
message = message & vbNewLine
MsgBox message

'polyA() = poly1() * poly2()
Call PolynomialMultiply(Poly1(), Poly2(), polyA())
message = "polyA() = poly1() * poly2()" & vbNewLine
message = message & "poly1:" & vbNewLine
message = message & ShowPoly(Poly1())
message = message & vbNewLine
message = message & "poly2:" & vbNewLine
message = message & ShowPoly(Poly2())
message = message & vbNewLine
message = message & "polyA" & vbNewLine
message = message & ShowPoly(polyA())
message = message & vbNewLine
MsgBox message

'Poly1() / Poly2() = polyA() + polyB() / Poly2()
Call PolynomialDivide(Poly1(), Poly2(), polyA(), polyB())
message = "Poly1() / Poly2() = polyA() + polyB() / Poly2()" & vbNewLine
message = message & "poly1:" & vbNewLine
message = message & ShowPoly(Poly1())
message = message & vbNewLine
message = message & "poly2:" & vbNewLine
message = message & ShowPoly(Poly2())
message = message & vbNewLine
message = message & "polyA" & vbNewLine
message = message & ShowPoly(polyA())
message = message & vbNewLine
message = message & "polyB" & vbNewLine
message = message & ShowPoly(polyB())
message = message & vbNewLine
MsgBox message

'PolynomialValue
message = "PolynomialValue" & vbNewLine
message = message & "poly1:" & vbNewLine
message = message & ShowPoly(Poly1())
message = message & vbNewLine
x = 0
message = message & "poly1(" & x & ") = " & _
          PolynomialValue(Poly1(), x) & vbNewLine
x = 1
message = message & "poly1(" & x & ") = " & _
          PolynomialValue(Poly1(), x) & vbNewLine
x = -1
message = message & "poly1(" & x & ") = " & _
          PolynomialValue(Poly1(), x) & vbNewLine
MsgBox message

End Sub
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub TestBairstowLaguerre()
Dim poly() As Double
Dim tmpPoly1(0 To 2) As Double, tmpPoly2() As Double
Dim RealRoot() As Double, ImagRoot() As Double
Dim TrueRealRoot() As Double, TrueImagRoot() As Double
Dim TimerStart As Date, TimerEnd As Date
Dim i As Integer, j As Integer, k As Integer, l As Integer
Const n As Integer = 20
Const rLM1 As Integer = -41
Const iLM1 As Integer = -43
Const trLM1 As Integer = -53
Const tiLM1 As Integer = -59
Dim message As String

ReDim RealRoot(1 + rLM1 To n + rLM1)
ReDim ImagRoot(1 + iLM1 To n + iLM1)
ReDim TrueRealRoot(1 + trLM1 To n + trLM1)
ReDim TrueImagRoot(1 + tiLM1 To n + tiLM1)
ReDim tmpPoly2(0 To n)
tmpPoly2(0) = 1#
For i = 1 To n
    tmpPoly2(i) = 0#
Next i
ReDim poly(0 To n)
For i = 0 To n
    poly(i) = 0#
Next i
Call Randomize
'Wilkinson's polynomial
'p(x) = (x - 1) * (x - 2) * (x - 3) * ... * (x - N)
'The problem of finding the roots is ill-conditioned:
'A small change in one coefficient can lead to drastic
'changes in the roots.
'If the magnitude of the coefficient of each term differs
'from each other drastically, the problem of finding the
'roots is quite ill-conditioned.
Dim HasConjugate As Boolean
HasConjugate = False
For i = 1 To n
'For non-ill-conditioned polynomial
    TrueRealRoot(i + trLM1) = 5 * (Rnd() - 0.5)
    TrueImagRoot(i + tiLM1) = 0#
    tmpPoly1(0) = -TrueRealRoot(i + trLM1)
    tmpPoly1(1) = 1
    tmpPoly1(2) = 0

'For ill-conditioned polynomial
'    TrueRealRoot(i + trLM1) = 3 * (Rnd() - 0.5)
'    TrueImagRoot(i + tiLM1) = 0#
'    tmpPoly1(0) = -TrueRealRoot(i + trLM1)
'    tmpPoly1(1) = 1
'    tmpPoly1(2) = 0

'For repeated roots
'    TrueRealRoot(i + trLM1) = Int(5 * 5 * (Rnd() - 0.5)) / 5.
'    TrueImagRoot(i + tiLM1) = 0#
'    tmpPoly1(0) = -TrueRealRoot(i + trLM1)
'    tmpPoly1(1) = 1
'    tmpPoly1(2) = 0

'For complex roots
'    If HasConjugate Then
'        TrueRealRoot(i + trLM1) = TrueRealRoot(i - 1 + trLM1)
'        TrueImagRoot(i + tiLM1) = -TrueImagRoot(i - 1 + tiLM1)
'        tmpPoly1(0) = TrueRealRoot(i + trLM1) ^ 2 + _
'                      TrueImagRoot(i + tiLM1) ^ 2
'        tmpPoly1(1) = -2# * TrueRealRoot(i + trLM1)
'        tmpPoly1(2) = 1
'    Else
'        TrueRealRoot(i + trLM1) = 5 * (Rnd() - 0.5)
'        TrueImagRoot(i + tiLM1) = 5 * (Rnd() - 0.5)
'        tmpPoly1(0) = 1
'        tmpPoly1(1) = 0
'        tmpPoly1(2) = 0
'    End If
'    HasConjugate = Not HasConjugate

'For complex roots,ill-conditioned polynomial
'    If HasConjugate Then
'        TrueRealRoot(i + trLM1) = TrueRealRoot(i - 1 + trLM1)
'        TrueImagRoot(i + tiLM1) = -TrueImagRoot(i - 1 + tiLM1)
'        tmpPoly1(0) = TrueRealRoot(i + trLM1) ^ 2 + _
'                      TrueImagRoot(i + tiLM1) ^ 2
'        tmpPoly1(1) = -2# * TrueRealRoot(i + trLM1)
'        tmpPoly1(2) = 1
'    Else
'        TrueRealRoot(i + trLM1) = 1.5 * (Rnd() - 0.5)
'        TrueImagRoot(i + tiLM1) = 1.5 * (Rnd() - 0.5)
'        tmpPoly1(0) = 1
'        tmpPoly1(1) = 0
'        tmpPoly1(2) = 0
'    End If
'    HasConjugate = Not HasConjugate

'For repeated, complex roots
'    If HasConjugate Then
'        TrueRealRoot(i + trLM1) = TrueRealRoot(i - 1 + trLM1)
'        TrueImagRoot(i + tiLM1) = -TrueImagRoot(i - 1 + tiLM1)
'        tmpPoly1(0) = TrueRealRoot(i + trLM1) ^ 2 + _
'                      TrueImagRoot(i + tiLM1) ^ 2
'        tmpPoly1(1) = -2# * TrueRealRoot(i + trLM1)
'        tmpPoly1(2) = 1
'    Else
'        TrueRealRoot(i + trLM1) = Int(7 * (Rnd() - 0.5)) / 5#
'        TrueImagRoot(i + tiLM1) = Int(7 * (Rnd() - 0.5)) / 5#
'        tmpPoly1(0) = 1
'        tmpPoly1(1) = 0
'        tmpPoly1(2) = 0
'    End If
'    HasConjugate = Not HasConjugate
    
    Call PolynomialMultiply(tmpPoly1(), tmpPoly2(), poly())
    j = 0
    k = i
    For l = j To k
        tmpPoly2(l) = poly(l)
    Next l
Next i
message = ""
message = message & "poly: " & vbNewLine
message = message & ShowPoly(poly()) & vbNewLine
MsgBox message

'less efficient
'less accurate
'worse for ill-conditioned polynomial
'fail for repeated roots
'Bairstow's method
message = ""
TimerStart = timer
For i = 1 To 100
    Call PolynomialRootBairstow(poly(), RealRoot(), ImagRoot())
Next i
TimerEnd = timer
message = message & "PolynomialRootBairstow: " & vbNewLine
message = message & "Time: " & _
          (TimerEnd - TimerStart) & " seconds" & vbNewLine
message = message & "RealRoot: " & vbNewLine
message = message & ShowPoly(RealRoot()) & vbNewLine
message = message & "ImagRoot: " & vbNewLine
message = message & ShowPoly(ImagRoot()) & vbNewLine
message = message & "PolynomialValue: " & vbNewLine
For i = 1 To n
    message = message & i & vbTab & PolynomialValue(poly(), RealRoot(i + rLM1)) & vbNewLine
Next i
MsgBox message

Worksheets("Bairstow").Activate
Columns(1).Clear
Columns(2).Clear
Columns(3).Clear
Columns(4).Clear
Cells(1, 1) = "Bairstow's method"
Cells(2, 1) = "TrueRealRoot"
Cells(2, 2) = "TrueImagRoot"
Cells(2, 3) = "RealRoot"
Cells(2, 4) = "ImagRoot"
Call QuickSort(RealRoot())
Call QuickSort(TrueRealRoot())
Call QuickSort(ImagRoot())
Call QuickSort(TrueImagRoot())
For i = 1 To n
    Cells(i + 2, 1).Value = TrueRealRoot(i + trLM1)
    Cells(i + 2, 2).Value = TrueImagRoot(i + tiLM1)
    Cells(i + 2, 3).Value = RealRoot(i + rLM1)
    Cells(i + 2, 4).Value = ImagRoot(i + iLM1)
Next i

'more efficient
'more accurate
'bettter for ill-conditioned polynomial
'work well for repeated roots
'Laguerre's method
message = ""
TimerStart = timer
For i = 1 To 100
    Call PolynomialRootLaguerre(poly(), RealRoot(), ImagRoot())
Next i
TimerEnd = timer
message = message & "PolynomialRootLaguerre: " & vbNewLine
message = message & "Time: " & _
          (TimerEnd - TimerStart) & " seconds" & vbNewLine
message = message & "RealRoot: " & vbNewLine
message = message & ShowPoly(RealRoot()) & vbNewLine
message = message & "ImagRoot: " & vbNewLine
message = message & ShowPoly(ImagRoot()) & vbNewLine
message = message & "PolynomialValue: " & vbNewLine
For i = 1 To n
    message = message & i & vbTab & PolynomialValue(poly(), RealRoot(i + rLM1)) & vbNewLine
Next i

MsgBox message

Worksheets("Laguerre").Activate
Columns(1).Clear
Columns(2).Clear
Columns(3).Clear
Columns(4).Clear
Cells(1, 1) = "Laguerre's method"
Cells(2, 1) = "TrueRealRoot"
Cells(2, 2) = "TrueImagRoot"
Cells(2, 3) = "RealRoot"
Cells(2, 4) = "ImagRoot"
Call QuickSort(ImagRoot())
Call QuickSort(TrueImagRoot())
Call QuickSort(RealRoot())
Call QuickSort(TrueRealRoot())
For i = 1 To n
    Cells(i + 2, 1).Value = TrueRealRoot(i + trLM1)
    Cells(i + 2, 2).Value = TrueImagRoot(i + tiLM1)
    Cells(i + 2, 3).Value = RealRoot(i + rLM1)
    Cells(i + 2, 4).Value = ImagRoot(i + iLM1)
Next i

End Sub
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Function PolynomialValue(poly() As Double, _
                                x As Double) As Double
Dim UpperBound As Integer
Dim LowerBound As Integer
Dim i As Integer

LowerBound = LBound(poly)
If LowerBound < 0 Then
    MsgBox "The lower bound of parameter Poly() is less than 0!"
    Exit Function
End If

UpperBound = UBound(poly)
For i = UpperBound To LowerBound Step -1
    If poly(i) <> 0# Then Exit For
Next i
If i < LowerBound Then
    PolynomialValue = 0#
    Exit Function
Else
    UpperBound = i
    For i = LowerBound To UpperBound
        If poly(i) <> 0# Then Exit For
    Next i
    LowerBound = i
End If

PolynomialValue = poly(UpperBound)
For i = UpperBound - 1 To LowerBound Step -1
    PolynomialValue = PolynomialValue * x + poly(i)
Next i

If LowerBound >= 40 Then
    PolynomialValue = PolynomialValue * x ^ LowerBound
Else
    For i = 1 To LowerBound
        PolynomialValue = PolynomialValue * x
    Next i
End If

End Function
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub PolynomialComplexValue(poly() As Double, _
           xReal As Double, xImaginary As Double, _
           PolyReal As Double, PolyImaginary As Double)
Dim UpperBound As Integer, LowerBound As Integer
Dim i As Integer
Dim tmpReal As Double, tmpImag As Double

LowerBound = LBound(poly)
If LowerBound < 0 Then
    MsgBox "The lower bound of parameter Poly() is less than 0!"
    Exit Sub
End If

UpperBound = UBound(poly)
For i = UpperBound To LowerBound Step -1
    If poly(i) <> 0# Then Exit For
Next i
If i < LowerBound Then
    PolyReal = 0#
    PolyImaginary = 0#
    Exit Sub
Else
    UpperBound = i
    For i = LowerBound To UpperBound
        If poly(i) <> 0# Then Exit For
    Next i
    LowerBound = i
End If

PolyReal = poly(UpperBound)
PolyImaginary = 0#
For i = UpperBound - 1 To LowerBound Step -1
    tmpReal = PolyReal
    tmpImag = PolyImaginary
    PolyReal = tmpReal * xReal - tmpImag * xImaginary + poly(i)
    PolyImaginary = tmpReal * xImaginary + tmpImag * xReal
Next i

If LowerBound >= 10 Then
    Dim r As Double, theta As Double
    Dim CosTheta As Double, SinTheta As Double
    r = Sqr(xReal * xReal + xImaginary * xImaginary)
    r = r ^ LowerBound
    theta = LowerBound * Atn2(xReal, xImaginary)
    CosTheta = Cos(theta)
    SinTheta = Sin(theta)
    tmpReal = PolyReal
    tmpImag = PolyImaginary
    PolyReal = r * (tmpReal * CosTheta - tmpImag * SinTheta)
    PolyImaginary = r * (tmpReal * SinTheta + tmpImag * CosTheta)
Else
    For i = 1 To LowerBound
        tmpReal = PolyReal
        tmpImag = PolyImaginary
        PolyReal = tmpReal * xReal - tmpImag * xImaginary
        PolyImaginary = tmpReal * xImaginary + tmpImag * xReal
    Next i
End If

End Sub
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
'polyDiff = d(poly)/dx = polyDiff(L-1 to U-1) (L-1) >= 0
'=L*poly(L)*x^(L-1) + (L+1)*poly(L+1)*x^(L) + ... + U*poly(U)*x^(U-1)
Public Sub PolynomialDifferentiate(poly() As Double, _
                                   polyDiff() As Double)
Dim UpperBound As Integer
Dim LowerBound As Integer
Dim polyDiffL As Integer
Dim polyDiffU As Integer
Dim i As Integer
Dim j As Integer

'UpperBound and LowerBound now is the upper bound and
'lower bound of the index of array poly
LowerBound = LBound(poly)
If LowerBound < 0 Then
    MsgBox "The lower bound of parameter Poly() is less than 0!"
    Exit Sub
End If

UpperBound = UBound(poly)
For i = UpperBound To LowerBound Step -1
    If poly(i) <> 0# Then Exit For
Next i
If i < LowerBound Then
    For j = LBound(polyDiff) To UBound(polyDiff)
        polyDiff(j) = 0#
    Next j
    Exit Sub
Else
    UpperBound = i
    For i = LowerBound To UpperBound
        If poly(i) <> 0# Then Exit For
    Next i
    LowerBound = i
End If

'UpperBound and LowerBound will be the upper bound and
'lower bound of the index of array polyDiff
UpperBound = UpperBound - 1
If LowerBound = 0 Then
    LowerBound = 0
Else
    LowerBound = LowerBound - 1
End If

polyDiffL = LBound(polyDiff)
polyDiffU = UBound(polyDiff)
If (polyDiffU >= UpperBound) And (polyDiffL <= LowerBound) Then
    If polyDiffL < 0 Then
        ReDim polyDiff(LowerBound To UpperBound)
        polyDiffL = LowerBound
        polyDiffU = UpperBound
    End If
Else
    ReDim polyDiff(LowerBound To UpperBound)
    polyDiffL = LowerBound
    polyDiffU = UpperBound
End If

For i = polyDiffL To LowerBound - 1
    polyDiff(i) = 0#
Next i
For i = LowerBound To UpperBound
    j = i + 1
    polyDiff(i) = j * poly(j)
Next i
For i = UpperBound + 1 To polyDiffU
    polyDiff(i) = 0#
Next i
End Sub
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
'polyInt = d(poly)/dx = polyInt(L+1 to U+1) (L-1) >= 0
'=poly(L)/(L+1)*x^(L+1) + poly(L+1)/(L+2)*x^(L+2) + ... + poly(U)/(U+1)*x^(U+1)
Public Sub PolynomialIntegrate(poly() As Double, _
                               polyIntegrate() As Double, _
                               Optional C0 As Double = 0#)
Dim UpperBound As Integer
Dim LowerBound As Integer
Dim polyIntegrateL As Integer
Dim polyIntegrateU As Integer
Dim i As Integer
Dim j As Integer

'UpperBound and LowerBound now is the upper bound and
'lower bound of the index of array poly
LowerBound = LBound(poly)
If LowerBound < 0 Then
    MsgBox "The lower bound of parameter Poly() is less than 0!"
    Exit Sub
End If

UpperBound = UBound(poly)
For i = UpperBound To LowerBound Step -1
    If poly(i) <> 0# Then Exit For
Next i
If i < LowerBound Then
    If LBound(polyIntegrate) <> 0 Then
        ReDim polyIntegrate(0 To 0)
    End If
    polyIntegrate(0) = C0
    For j = 1 To UBound(polyIntegrate)
        polyIntegrate(j) = 0#
    Next j
    Exit Sub
Else
    UpperBound = i
    For i = LowerBound To UpperBound
        If poly(i) <> 0# Then Exit For
    Next i
    LowerBound = i
End If

polyIntegrateL = LBound(polyIntegrate)
polyIntegrateU = UBound(polyIntegrate)
If (polyIntegrateU <= UpperBound) Or _
   (polyIntegrateL <> 0) Then
    ReDim polyIntegrate(0 To UpperBound + 1)
    polyIntegrateL = 0
    polyIntegrateU = UpperBound + 1
End If

polyIntegrate(0) = C0
For i = 1 To LowerBound
    polyIntegrate(i) = 0#
Next i
For i = LowerBound To UpperBound
    j = i + 1
    polyIntegrate(j) = poly(i) / CDbl(j)
Next i
For i = UpperBound + 2 To polyIntegrateU
    polyIntegrate(i) = 0#
Next i
End Sub
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
'poly = (x^2 + C1*x + C0) * Quotient + (R1*x + R0)
'
'R1 + d(R1)/d(C1) * dC1 + d(R1)/d(C0) * dC0 = 0
'R0 + d(R0)/d(C1) * dC1 + d(R0)/d(C0) * dC0 = 0
'update C1 & C0 with (C1 + dC1) & (C0 + dC0)
Private Sub PolynomialRootBairstow(poly() As Double, _
                                   RealRoot() As Double, _
                                   ImaginaryRoot() As Double)
Dim c1 As Double, C0 As Double
Dim c2 As Double, C2Inv As Double
Dim RealRoot1 As Double, ImagRoot1 As Double
Dim RealRoot2 As Double, ImagRoot2 As Double
Dim R1 As Double, R0 As Double
Dim dR1dC1 As Double, dR1dC0 As Double
Dim dR0dC1 As Double, dR0dC0 As Double
Dim tmp As Double, tmp1 As Double, tmp2 As Double
Dim residue1 As Double, residue2 As Double
Const dC1 As Double = 0.0000001
Const dC0 As Double = 0.0000001
Const dC1inv As Double = 1# / dC1
Const dC0inv As Double = 1# / dC0
Const tolerence As Double = 0.000000001
Dim Quadratic(0 To 2) As Double
Dim Remainder(0 To 1) As Double
Dim tmpPoly() As Double
Dim Quotient() As Double
Dim tmpQuotient() As Double
Dim UpperBound As Integer
Dim RootIndex As Integer
Dim rLM1 As Integer, iLM1 As Integer
Dim i As Integer, j As Integer, k As Integer

j = LBound(poly)
k = UBound(poly)
If j < 0 Then
    MsgBox "The lower bound of parameter Poly() is less than 0!"
    Exit Sub
End If

For i = k To j Step -1
    If poly(i) <> 0 Then
        Exit For
    End If
Next i
If i < j Then
    MsgBox "The values of all the elements of parameter array poly() are all zero!"
    Exit Sub
Else
    UpperBound = i
    For i = j To UpperBound
        If poly(i) <> 0 Then
            Exit For
        End If
    Next i
    j = i
End If

tmp = 1# / Log(10)
'tmp1 stands for max order of magnitude of coefficients
tmp1 = Log(Abs(poly(j))) * tmp
'tmp2 stands for min order of magnitude of coefficients
tmp2 = Log(Abs(poly(j))) * tmp
For i = j + 1 To UpperBound
    c2 = Abs(poly(i))
    If c2 <> 0# Then
        C2Inv = Log(Abs(poly(i))) * tmp
        If C2Inv > tmp1 Then tmp1 = C2Inv
        If C2Inv < tmp2 Then tmp2 = C2Inv
    End If
Next i
If (tmp1 - tmp2) > 6# Then
    MsgBox "The magnitude of coefficients of the polynomial" & vbNewLine & _
           "differs from each other drastically." & vbNewLine & _
           "Finding the roots may be quite ill-conditioned!"
End If

rLM1 = LBound(RealRoot) - 1
iLM1 = LBound(ImaginaryRoot) - 1
If (UBound(RealRoot) - rLM1) <> UpperBound Then
    ReDim RealRoot(1 To UpperBound)
    rLM1 = 0
End If
If (UBound(ImaginaryRoot) - iLM1) <> UpperBound Then
    ReDim ImaginaryRoot(1 To UpperBound)
    iLM1 = 0
End If
For i = 1 To UpperBound
    RealRoot(i + rLM1) = 0#
    ImaginaryRoot(i + iLM1) = 0#
Next i
'Normalize Polynomial
tmp = 1# / poly(UpperBound)
ReDim tmpPoly(0 To UpperBound)
For i = 0 To j - 1
    tmpPoly(i) = 0#
Next i
For i = j To UpperBound
    tmpPoly(i) = poly(i) * tmp
Next i

ReDim Quotient(0 To UpperBound - 2)
ReDim tmpQuotient(0 To UpperBound - 2)
For i = j To UpperBound - 2
    Quotient(i) = 0#
Next i
'Quadratic = 1*x^2 + C1*x + C0
Quadratic(2) = 1#
RootIndex = 0
Do While ((UpperBound - RootIndex) > 2)
    RealRoot1 = 0#
    ImagRoot1 = 0#
    RealRoot2 = 0#
    ImagRoot2 = 0#
    C0 = RealRoot1 * RealRoot2
    c1 = -(RealRoot1 + RealRoot2)
    Quadratic(1) = c1
    Quadratic(0) = C0
    Call PolynomialDivide(tmpPoly(), Quadratic(), Quotient(), Remainder())
    R1 = Remainder(1)
    R0 = Remainder(0)
    residue1 = Abs(R1 * RealRoot1 + R0)
    residue2 = Abs(R1 * RealRoot2 + R0)
    Do While (residue1 > tolerence Or residue2 > tolerence)
'dR1dC1 and dR0dC1
        Quadratic(1) = c1 + dC1
        Call PolynomialDivide(tmpPoly(), Quadratic(), tmpQuotient(), Remainder())
        dR1dC1 = (Remainder(1) - R1) * dC1inv
        dR0dC1 = (Remainder(0) - R0) * dC1inv
'dR1dC0 and dR0dC0
        Quadratic(1) = c1
        Quadratic(0) = C0 + dC0
        Call PolynomialDivide(tmpPoly(), Quadratic(), tmpQuotient(), Remainder())
        dR1dC0 = (Remainder(1) - R1) * dC0inv
        dR0dC0 = (Remainder(0) - R0) * dC0inv
'Update C1 and C0
        tmp = 1# / (dR1dC1 * dR0dC0 - dR1dC0 * dR0dC1)
        tmp1 = dR1dC0 * R0 - R1 * dR0dC0
        tmp2 = R1 * dR0dC1 - dR1dC1 * R0
        c1 = c1 + tmp1 * tmp
        C0 = C0 + tmp2 * tmp
'Remainder for new C1 and C0
        Quadratic(1) = c1
        Quadratic(0) = C0
        Call PolynomialDivide(tmpPoly(), Quadratic(), Quotient(), Remainder())
        R1 = Remainder(1)
        R0 = Remainder(0)
'Roots for 1*x^2 + C1*x + C0 and Polynomial value for roots
        tmp = c1 ^ 2 - 4# * C0
        If (tmp >= 0) Then
            tmp1 = -0.5 * c1
            tmp2 = 0.5 * Sqr(tmp)
            RealRoot1 = tmp1 - tmp2
            ImagRoot1 = 0#
            RealRoot2 = tmp1 + tmp2
            ImagRoot2 = 0#
            residue1 = Abs(R1 * RealRoot1 + R0)
            residue2 = Abs(R1 * RealRoot2 + R0)
        Else
            tmp = -tmp
            tmp1 = -0.5 * c1
            tmp2 = 0.5 * Sqr(tmp)
            RealRoot1 = tmp1
            ImagRoot1 = tmp2
            RealRoot2 = tmp1
            ImagRoot2 = -tmp2
            residue1 = Abs(R1 * RealRoot1 + R0)
            residue2 = Abs(R1 * ImagRoot1)
        End If
    Loop
'Roots for 1*x^2 + C1*x + C0
    RootIndex = RootIndex + 1
    RealRoot(RootIndex + rLM1) = RealRoot1
    ImaginaryRoot(RootIndex + iLM1) = ImagRoot1
    RootIndex = RootIndex + 1
    RealRoot(RootIndex + rLM1) = RealRoot2
    ImaginaryRoot(RootIndex + iLM1) = ImagRoot2
    j = UpperBound - RootIndex
    For i = 0 To j
        tmpPoly(i) = Quotient(i)
    Next i
    tmpPoly(j + 1) = 0#
    tmpPoly(j + 2) = 0#
Loop

If (UpperBound - RootIndex) = 1 Then
    RootIndex = RootIndex + 1
    RealRoot(RootIndex + rLM1) = -tmpPoly(0) / tmpPoly(1)
    ImaginaryRoot(RootIndex + iLM1) = 0#
ElseIf (UpperBound - RootIndex) = 2 Then
    c2 = tmpPoly(2)
    c1 = tmpPoly(1)
    C0 = tmpPoly(0)
    C2Inv = 1# / c2
    tmp = c1 ^ 2 - 4# * c2 * C0
    If (tmp >= 0) Then
        R0 = -0.5 * c1 * C2Inv
        R1 = 0.5 * Sqr(tmp) * C2Inv
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = R0 - R1
        ImaginaryRoot(RootIndex + iLM1) = 0#
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = R0 + R1
        ImaginaryRoot(RootIndex + iLM1) = 0#
    Else
        tmp = -tmp
        R0 = -0.5 * c1 * C2Inv
        R1 = 0.5 * Sqr(tmp) * C2Inv
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = R0
        ImaginaryRoot(RootIndex + iLM1) = R1
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = R0
        ImaginaryRoot(RootIndex + iLM1) = -R1
    End If
End If

End Sub
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
'denote d(poly)/dx by dpoly, d2(poly)/dx2 by ddpoly
'
'update root x with x - n * poly / (dpoly +- sqr(h))
'h = (n-1) * ((n-1) * dpoly^2 - n * poly * ddpoly)
'the sign in the denominator is determined so that
'it makes the denominator to be larger.
Public Sub PolynomialRootLaguerre(poly() As Double, RealRoot() As Double, ImaginaryRoot() As Double)
Dim tmpPoly() As Double, dtmpPoly() As Double, ddtmpPoly() As Double
Dim Quadratic(0 To 2) As Double
Dim Remainder(0 To 1) As Double
Dim Quotient() As Double
Dim residue As Double
Dim tmp As Double, tmp1 As Double, tmp2 As Double
Dim tmpReal As Double, tmpImag As Double
Dim dtmpReal As Double, dtmpImag As Double
Dim ddtmpReal As Double, ddtmpImag As Double
Dim tmp1Real As Double, tmp1Imag As Double
Dim tmp2Real As Double, tmp2Imag As Double
Dim RealRootNew As Double, ImagRootNew As Double
Dim RealRootOld As Double, ImagRootOld As Double
Dim Realh As Double, Imagh As Double
Dim c2 As Double, c1 As Double, C0 As Double, C2Inv As Double
Const tolerence As Double = 0.000000001
Dim i As Integer, j As Integer, k As Integer
Dim RootIndex As Integer
Dim UpperBound As Integer, UpperBoundM1 As Integer
Dim rLM1 As Integer, iLM1 As Integer

j = LBound(poly)
k = UBound(poly)
If j < 0 Then
    MsgBox "The lower bound of parameter Poly() is less than 0!"
    Exit Sub
End If

For i = k To j Step -1
    If poly(i) <> 0 Then
        Exit For
    End If
Next i
If i < j Then
    MsgBox "The values of all the elements of parameter array poly() are all zero!"
    Exit Sub
Else
    UpperBound = i
    For i = j To UpperBound
        If poly(i) <> 0 Then
            Exit For
        End If
    Next i
    j = i
End If

tmp = 1# / Log(10)
'tmp1 stands for max order of magnitude of coefficients
tmp1 = -1000#
'tmp2 stands for min order of magnitude of coefficients
tmp2 = 1000#
For i = j To k
    c2 = Abs(poly(i))
    If c2 <> 0# Then
        C2Inv = Log(Abs(poly(i))) * tmp
        If C2Inv > tmp1 Then tmp1 = C2Inv
        If C2Inv < tmp2 Then tmp2 = C2Inv
    End If
Next i
If (tmp1 - tmp2) > 6# Then
    MsgBox "The magnitude of coefficients of the polynomial" & vbNewLine & _
           "differs from each other drastically." & vbNewLine & _
           "Finding the roots may be quite ill-conditioned!"
End If

rLM1 = LBound(RealRoot) - 1
iLM1 = LBound(ImaginaryRoot) - 1
If (UBound(RealRoot) - rLM1) <> UpperBound Then
    ReDim RealRoot(1 To UpperBound)
    rLM1 = 0
End If
If (UBound(ImaginaryRoot) - iLM1) <> UpperBound Then
    ReDim ImaginaryRoot(1 To UpperBound)
    iLM1 = 0
End If
For i = 1 To UpperBound
    RealRoot(i + rLM1) = 0#
    ImaginaryRoot(i + iLM1) = 0#
Next i
'Normalize Polynomial
tmp = 1# / poly(UpperBound)
ReDim tmpPoly(0 To UpperBound)
For i = 0 To j - 1
    tmpPoly(i) = 0#
Next i
For i = j To UpperBound
    tmpPoly(i) = poly(i) * tmp
Next i

ReDim Quotient(0 To UpperBound - 2)
ReDim dtmpPoly(0 To UpperBound - 1)
ReDim ddtmpPoly(0 To UpperBound - 2)

Quadratic(2) = 1#
RootIndex = 0
Do While (UpperBound > 2)
    UpperBoundM1 = UpperBound - 1
    Call PolynomialDifferentiate(tmpPoly(), dtmpPoly())
    Call PolynomialDifferentiate(dtmpPoly(), ddtmpPoly())
    
    RealRootNew = 0#
    ImagRootNew = 0#
    residue = Abs(PolynomialValue(tmpPoly(), RealRootNew))
    Do While (residue > tolerence)
        If ImagRootNew = 0# Then
            RealRootOld = RealRootNew
            ImagRootOld = ImagRootNew
            tmpReal = PolynomialValue(tmpPoly(), RealRootOld)
            dtmpReal = PolynomialValue(dtmpPoly(), RealRootOld)
            ddtmpReal = PolynomialValue(ddtmpPoly(), RealRootOld)
            Realh = UpperBoundM1 * (UpperBoundM1 * dtmpReal ^ 2 - _
                                    UpperBound * tmpReal * ddtmpReal)
            If Realh < 0# Then
                Realh = -Realh
                tmp1Real = UpperBound * tmpReal / (dtmpReal ^ 2 + Realh)
                RealRootNew = RealRootOld - tmp1Real * dtmpReal
                ImagRootNew = tmp1Real * Sqr(Realh)
                Call PolynomialComplexValue(tmpPoly(), RealRootNew, ImagRootNew, _
                                            tmp1Real, tmp1Imag)
                residue = Sqr(tmp1Real ^ 2 + tmp1Imag ^ 2)
            Else
                RealRootNew = RealRootOld - UpperBound * tmpReal / _
                              (dtmpReal + Sgn(dtmpReal) * Sqr(Realh))
                ImagRootNew = 0#
                residue = Abs(PolynomialValue(tmpPoly(), RealRootNew))
            End If
        Else
            RealRootOld = RealRootNew
            ImagRootOld = ImagRootNew
            Call PolynomialComplexValue(tmpPoly(), _
                                        RealRootOld, ImagRootOld, _
                                        tmpReal, tmpImag)
            Call PolynomialComplexValue(dtmpPoly(), _
                                        RealRootOld, ImagRootOld, _
                                        dtmpReal, dtmpImag)
            Call PolynomialComplexValue(ddtmpPoly(), _
                                        RealRootOld, ImagRootOld, _
                                        ddtmpReal, ddtmpImag)
            tmp1Real = UpperBoundM1 * (dtmpReal ^ 2 - dtmpImag ^ 2)
            tmp1Imag = UpperBoundM1 * 2# * dtmpReal * dtmpImag
            tmp2Real = UpperBound * (tmpReal * ddtmpReal - tmpImag * ddtmpImag)
            tmp2Imag = UpperBound * (tmpReal * ddtmpImag + ddtmpReal * tmpImag)
            Realh = UpperBoundM1 * (tmp1Real - tmp2Real)
            Imagh = UpperBoundM1 * (tmp1Imag - tmp2Imag)
            tmp = Sqr(Sqr(Realh ^ 2 + Imagh ^ 2))
            tmp1 = 0.5 * Atn2(Realh, Imagh)
            tmp1Real = tmp * Cos(tmp1)
            tmp1Imag = tmp * Sin(tmp1)
            If (dtmpReal * tmp1Real + dtmpImag * tmp1Imag) > 0# Then
                tmp2Real = dtmpReal + tmp1Real
                tmp2Imag = dtmpImag + tmp1Imag
            Else
                tmp2Real = dtmpReal - tmp1Real
                tmp2Imag = dtmpImag - tmp1Imag
            End If
            tmp = 1# / (tmp2Real ^ 2 + tmp2Imag ^ 2)
            RealRootNew = RealRootOld - UpperBound * _
                          (tmpReal * tmp2Real + tmpImag * tmp2Imag) * tmp
            ImagRootNew = ImagRootOld - UpperBound * _
                          (tmp2Real * tmpImag - tmpReal * tmp2Imag) * tmp
            Call PolynomialComplexValue(tmpPoly(), RealRootNew, ImagRootNew, _
                                        tmp1Real, tmp1Imag)
            residue = Sqr(tmp1Real ^ 2 + tmp1Imag ^ 2)
        End If
    Loop
    If ImagRootNew = 0# Then
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = RealRootNew
        ImaginaryRoot(RootIndex + iLM1) = 0#
        dtmpPoly(UpperBoundM1) = tmpPoly(UpperBound)
        For i = UpperBound - 2 To 0 Step -1
            k = i + 1
            dtmpPoly(i) = tmpPoly(k) + RealRootNew * dtmpPoly(k)
        Next i
        For i = 0 To UpperBoundM1
            tmpPoly(i) = dtmpPoly(i)
        Next i
        tmpPoly(UpperBound) = 0#
        UpperBound = UpperBoundM1
    Else
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = RealRootNew
        ImaginaryRoot(RootIndex + iLM1) = ImagRootNew
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = RealRootNew
        ImaginaryRoot(RootIndex + iLM1) = -ImagRootNew
        Quadratic(1) = -2# * RealRootNew
        Quadratic(0) = RealRootNew ^ 2 + ImagRootNew ^ 2
        Call PolynomialDivide(tmpPoly(), Quadratic(), Quotient(), Remainder())
        j = UpperBound - 2
        For i = 0 To j
            tmpPoly(i) = Quotient(i)
        Next i
        tmpPoly(j + 1) = 0#
        tmpPoly(j + 2) = 0#
        UpperBound = j
    End If
Loop

If (UpperBound = 1) Then
    RootIndex = RootIndex + 1
    RealRoot(RootIndex + rLM1) = -tmpPoly(0) / tmpPoly(1)
    ImaginaryRoot(RootIndex + iLM1) = 0#
ElseIf (UpperBound = 2) Then
    c2 = tmpPoly(2)
    c1 = tmpPoly(1)
    C0 = tmpPoly(0)
    C2Inv = 1# / c2
    tmp = c1 ^ 2 - 4# * c2 * C0
    If (tmp >= 0) Then
        tmp1Real = -0.5 * c1 * C2Inv
        tmp2Real = 0.5 * Sqr(tmp) * C2Inv
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = tmp1Real - tmp2Real
        ImaginaryRoot(RootIndex + iLM1) = 0#
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = tmp1Real + tmp2Real
        ImaginaryRoot(RootIndex + iLM1) = 0#
    Else
        tmp = -tmp
        tmpReal = -0.5 * c1 * C2Inv
        tmpImag = 0.5 * Sqr(tmp) * C2Inv
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = tmpReal
        ImaginaryRoot(RootIndex + iLM1) = tmpImag
        RootIndex = RootIndex + 1
        RealRoot(RootIndex + rLM1) = tmpReal
        ImaginaryRoot(RootIndex + iLM1) = -tmpImag
    End If
End If

End Sub
'PolyResult = Poly1 + Poly2
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub PolynomialAdd(Poly1() As Double, _
                         Poly2() As Double, _
                         PolyResult() As Double)
Dim UpperBound1 As Integer, LowerBound1 As Integer
Dim UpperBound2 As Integer, LowerBound2 As Integer
Dim UpperBoundResult As Integer, LowerBoundResult As Integer
Dim UpperBound As Integer, LowerBound As Integer
Dim UpperUpperBound As Integer, LowerLowerBound As Integer
Dim i As Integer
Dim IsUpperUpper1 As Boolean
Dim IsLowerLower1 As Boolean

UpperBound1 = UBound(Poly1)
LowerBound1 = LBound(Poly1)
UpperBound2 = UBound(Poly2)
LowerBound2 = LBound(Poly2)
UpperBoundResult = UBound(PolyResult)
LowerBoundResult = LBound(PolyResult)

If LowerBound1 < 0 Then
    MsgBox "The lower bound of parameter Poly1() is less than 0!"
    Exit Sub
End If
If LowerBound2 < 0 Then
    MsgBox "The lower bound of parameter Poly2() is less than 0!"
    Exit Sub
End If
If LowerBoundResult < 0 Then
    MsgBox "The lower bound of parameter PolyResult() is less than 0!"
    Exit Sub
End If

For i = UpperBound1 To LowerBound1 Step -1
    If Poly1(i) <> 0# Then Exit For
Next i
If i < LowerBound1 Then
    UpperBound1 = LowerBound1
Else
    UpperBound1 = i
    For i = LowerBound1 To UpperBound1
        If Poly1(i) <> 0# Then Exit For
    Next i
    LowerBound1 = i
End If

For i = UpperBound2 To LowerBound2 Step -1
    If Poly2(i) <> 0# Then Exit For
Next i
If i < LowerBound2 Then
    UpperBound2 = LowerBound2
Else
    UpperBound2 = i
    For i = LowerBound2 To UpperBound2
        If Poly2(i) <> 0# Then Exit For
    Next i
    LowerBound2 = i
End If

If LowerBound2 > LowerBound1 Then
    LowerBound = LowerBound2
    LowerLowerBound = LowerBound1
    IsLowerLower1 = True
Else
    LowerBound = LowerBound1
    LowerLowerBound = LowerBound2
    IsLowerLower1 = False
End If

If UpperBound2 > UpperBound1 Then
    UpperBound = UpperBound1
    UpperUpperBound = UpperBound2
    IsUpperUpper1 = False
Else
    UpperBound = UpperBound2
    UpperUpperBound = UpperBound1
    IsUpperUpper1 = True
End If

If LowerBoundResult > LowerLowerBound Or _
   UpperBoundResult < UpperUpperBound Then
    ReDim PolyResult(LowerLowerBound To UpperUpperBound)
    LowerBoundResult = LowerLowerBound
    UpperBoundResult = UpperUpperBound
End If

For i = LowerBoundResult To LowerLowerBound - 1
    PolyResult(i) = 0#
Next i
If IsLowerLower1 Then
    For i = LowerLowerBound To LowerBound - 1
        PolyResult(i) = Poly1(i)
    Next i
Else
    For i = LowerLowerBound To LowerBound - 1
        PolyResult(i) = Poly2(i)
    Next i
End If
For i = LowerBound To UpperBound
    PolyResult(i) = Poly1(i) + Poly2(i)
Next i
If IsUpperUpper1 Then
    For i = UpperBound + 1 To UpperUpperBound
        PolyResult(i) = Poly1(i)
    Next i
Else
    For i = UpperBound + 1 To UpperUpperBound
        PolyResult(i) = Poly2(i)
    Next i
End If
For i = UpperUpperBound + 1 To UpperBoundResult
    PolyResult(i) = 0#
Next i
End Sub
'PolyResult = PolyLeft - PolyRight
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub PolynomialSubtract(PolyLeft() As Double, _
                              PolyRight() As Double, _
                              PolyResult() As Double)
Dim UpperBoundLeft As Integer, LowerBoundLeft As Integer
Dim UpperBoundRight As Integer, LowerBoundRight As Integer
Dim UpperBoundResult As Integer, LowerBoundResult As Integer
Dim UpperBound As Integer, LowerBound As Integer
Dim UpperUpperBound As Integer, LowerLowerBound As Integer
Dim i As Integer
Dim IsUpperUpperLeft As Boolean
Dim IsLowerLowerLeft As Boolean

UpperBoundLeft = UBound(PolyLeft)
LowerBoundLeft = LBound(PolyLeft)
UpperBoundRight = UBound(PolyRight)
LowerBoundRight = LBound(PolyRight)
UpperBoundResult = UBound(PolyResult)
LowerBoundResult = LBound(PolyResult)

If LowerBoundLeft < 0 Then
    MsgBox "The lower bound of parameter PolyLeft() is less than 0!"
    Exit Sub
End If
If LowerBoundRight < 0 Then
    MsgBox "The lower bound of parameter PolyRight() is less than 0!"
    Exit Sub
End If
If LowerBoundResult < 0 Then
    MsgBox "The lower bound of parameter PolyResult() is less than 0!"
    Exit Sub
End If

For i = UpperBoundLeft To LowerBoundLeft Step -1
    If PolyLeft(i) <> 0# Then Exit For
Next i
If i < LowerBoundLeft Then
    UpperBoundLeft = LowerBoundLeft
Else
    UpperBoundLeft = i
    For i = LowerBoundLeft To UpperBoundLeft
        If PolyLeft(i) <> 0# Then Exit For
    Next i
    LowerBoundLeft = i
End If

For i = UpperBoundRight To LowerBoundRight Step -1
    If PolyRight(i) <> 0# Then Exit For
Next i
If i < LowerBoundRight Then
    UpperBoundRight = LowerBoundRight
Else
    UpperBoundRight = i
    For i = LowerBoundRight To UpperBoundRight
        If PolyRight(i) <> 0# Then Exit For
    Next i
    LowerBoundRight = i
End If

If LowerBoundRight > LowerBoundLeft Then
    LowerBound = LowerBoundRight
    LowerLowerBound = LowerBoundLeft
    IsLowerLowerLeft = True
Else
    LowerBound = LowerBoundLeft
    LowerLowerBound = LowerBoundRight
    IsLowerLowerLeft = False
End If

If UpperBoundRight > UpperBoundLeft Then
    UpperBound = UpperBoundLeft
    UpperUpperBound = UpperBoundRight
    IsUpperUpperLeft = False
Else
    UpperBound = UpperBoundRight
    UpperUpperBound = UpperBoundLeft
    IsUpperUpperLeft = True
End If

If LowerBoundResult > LowerLowerBound Or _
   UpperBoundResult < UpperUpperBound Then
    ReDim PolyResult(LowerLowerBound To UpperUpperBound)
    LowerBoundResult = LowerLowerBound
    UpperBoundResult = UpperUpperBound
End If

For i = LowerBoundResult To LowerLowerBound - 1
    PolyResult(i) = 0#
Next i
If IsLowerLowerLeft Then
    For i = LowerLowerBound To LowerBound - 1
        PolyResult(i) = PolyLeft(i)
    Next i
Else
    For i = LowerLowerBound To LowerBound - 1
        PolyResult(i) = -PolyRight(i)
    Next i
End If
For i = LowerBound To UpperBound
    PolyResult(i) = PolyLeft(i) - PolyRight(i)
Next i
If IsUpperUpperLeft Then
    For i = UpperBound + 1 To UpperUpperBound
        PolyResult(i) = PolyLeft(i)
    Next i
Else
    For i = UpperBound + 1 To UpperUpperBound
        PolyResult(i) = -PolyRight(i)
    Next i
End If
For i = UpperUpperBound + 1 To UpperBoundResult
    PolyResult(i) = 0#
Next i

End Sub
'PolyResult = Poly1 * Poly2
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub PolynomialMultiply(Poly1() As Double, _
                              Poly2() As Double, _
                              PolyResult() As Double)
Dim UpperBound1 As Integer, LowerBound1 As Integer
Dim UpperBound2 As Integer, LowerBound2 As Integer
Dim UpperBoundResult As Integer, LowerBoundResult As Integer
Dim UpperBound As Integer, LowerBound As Integer
Dim i As Integer, j As Integer, k As Integer
Dim tmp As Double

UpperBound1 = UBound(Poly1)
LowerBound1 = LBound(Poly1)
UpperBound2 = UBound(Poly2)
LowerBound2 = LBound(Poly2)
UpperBoundResult = UBound(PolyResult)
LowerBoundResult = LBound(PolyResult)

If LowerBound1 < 0 Then
    MsgBox "The lower bound of parameter Poly1() is less than 0!"
    Exit Sub
End If
If LowerBound2 < 0 Then
    MsgBox "The lower bound of parameter Poly2() is less than 0!"
    Exit Sub
End If
If LowerBoundResult < 0 Then
    MsgBox "The lower bound of parameter PolyResult() is less than 0!"
    Exit Sub
End If

For i = UpperBound1 To LowerBound1 Step -1
    If Poly1(i) <> 0# Then Exit For
Next i
If i < LowerBound1 Then
    For i = LBound(PolyResult) To UBound(PolyResult)
        PolyResult(i) = 0#
    Next i
    Exit Sub
Else
    UpperBound1 = i
    For i = LowerBound1 To UpperBound1
        If Poly1(i) <> 0# Then Exit For
    Next i
    LowerBound1 = i
End If

For i = UpperBound2 To LowerBound2 Step -1
    If Poly2(i) <> 0# Then Exit For
Next i
If i < LowerBound2 Then
    For i = LBound(PolyResult) To UBound(PolyResult)
        PolyResult(i) = 0#
    Next i
    Exit Sub
Else
    UpperBound2 = i
    For i = LowerBound2 To UpperBound2
        If Poly2(i) <> 0# Then Exit For
    Next i
    LowerBound2 = i
End If

LowerBound = LowerBound1 + LowerBound2
UpperBound = UpperBound1 + UpperBound2

If LowerBoundResult > LowerBound Or _
   UpperBoundResult < UpperBound Then
    ReDim PolyResult(LowerBound To UpperBound)
    LowerBoundResult = LowerBound
    UpperBoundResult = UpperBound
End If

For k = LowerBoundResult To UpperBoundResult
    PolyResult(k) = 0#
Next k

For i = LowerBound1 To UpperBound1
    tmp = Poly1(i)
    For j = LowerBound2 To UpperBound2
        k = i + j
        PolyResult(k) = PolyResult(k) + tmp * Poly2(j)
    Next j
Next i

End Sub
'Numerator / Denominator = Quotient + Remainder / Denominator
'poly(L to U) = real-coefficient polynomial =
'poly(L)*x^L + poly(L+1)*x^(L+1) + ... + poly(U)*x^U
Public Sub PolynomialDivide(PolyNumerator() As Double, PolyDenominator() As Double, _
                            PolyQuotient() As Double, PolyRemainder() As Double)
Dim UpperBoundNumerator As Integer, LowerBoundNumerator As Integer
Dim UpperBoundDenominator As Integer, LowerBoundDenominator As Integer
Dim UpperBoundQuotient As Integer, LowerBoundQuotient As Integer
Dim UpperBoundRemainder As Integer, LowerBoundRemainder As Integer
Dim i As Integer, j As Integer, k As Integer
Dim tmpRemainder() As Double
Dim InvUpperDenominator As Double
Dim ratio As Double

UpperBoundNumerator = UBound(PolyNumerator)
LowerBoundNumerator = LBound(PolyNumerator)
UpperBoundDenominator = UBound(PolyDenominator)
LowerBoundDenominator = LBound(PolyDenominator)

If LowerBoundNumerator < 0 Then
    MsgBox "The lower bound of parameter PolyNumerator() is less than 0!"
    Exit Sub
End If
If LowerBoundDenominator < 0 Then
    MsgBox "The lower bound of parameter PolyDenominator() is less than 0!"
    Exit Sub
End If
If LBound(PolyQuotient) < 0 Then
    MsgBox "The lower bound of parameter PolyQuotient() is less than 0!"
    Exit Sub
End If
If LBound(PolyRemainder) < 0 Then
    MsgBox "The lower bound of parameter PolyRemainder() is less than 0!"
    Exit Sub
End If

For i = UpperBoundNumerator To LowerBoundNumerator Step -1
    If PolyNumerator(i) <> 0# Then Exit For
Next i
If i < LowerBoundNumerator Then
    For i = LBound(PolyQuotient) To UBound(PolyQuotient)
        PolyQuotient(i) = 0#
    Next i
    For i = LBound(PolyRemainder) To UBound(PolyRemainder)
        PolyRemainder(i) = 0#
    Next i
    Exit Sub
Else
    UpperBoundNumerator = i
    For i = LowerBoundNumerator To UpperBoundNumerator
        If PolyNumerator(i) <> 0# Then Exit For
    Next i
    LowerBoundNumerator = i
End If

For i = UpperBoundDenominator To LowerBoundDenominator Step -1
    If PolyDenominator(i) <> 0# Then Exit For
Next i
If i < LowerBoundDenominator Then
    MsgBox "The values of all the elements of parameter array PolyDenominator() are all zero!"
    Exit Sub
Else
    UpperBoundDenominator = i
    For i = LowerBoundDenominator To UpperBoundDenominator
        If PolyDenominator(i) <> 0# Then Exit For
    Next i
    LowerBoundDenominator = i
End If

If UpperBoundDenominator > UpperBoundNumerator Then
    For i = LBound(PolyQuotient) To UBound(PolyQuotient)
        PolyQuotient(i) = 0#
    Next i
    UpperBoundRemainder = UBound(PolyRemainder)
    LowerBoundRemainder = LBound(PolyRemainder)
    If UpperBoundRemainder < UpperBoundNumerator Or _
       LowerBoundRemainder > LowerBoundNumerator Then
        ReDim PolyRemainder(LowerBoundNumerator To UpperBoundNumerator)
        UpperBoundRemainder = UpperBoundNumerator
        LowerBoundRemainder = UpperBoundNumerator
    End If
    For i = LowerBoundRemainder To LowerBoundNumerator - 1
        PolyRemainder(i) = 0#
    Next i
    For i = LowerBoundNumerator To UpperBoundNumerator
        PolyRemainder(i) = PolyNumerator(i)
    Next i
    For i = UpperBoundNumerator + 1 To UpperBoundRemainder
        PolyRemainder(i) = 0#
    Next i
Else
    UpperBoundQuotient = UpperBoundNumerator - UpperBoundDenominator
    If LBound(PolyQuotient) <> 0 Or _
       UBound(PolyQuotient) < UpperBoundQuotient Then
        ReDim PolyQuotient(0 To UpperBoundQuotient)
    End If
    UpperBoundRemainder = UpperBoundDenominator - 1
    If LBound(PolyRemainder) <> 0 Or _
       UBound(PolyRemainder) < UpperBoundRemainder Then
        ReDim PolyRemainder(0 To UpperBoundRemainder)
    End If
    
    ReDim tmpRemainder(0 To UpperBoundNumerator)
    For i = 0 To LowerBoundNumerator - 1
        tmpRemainder(i) = 0#
    Next i
    For i = LowerBoundNumerator To UpperBoundNumerator
        tmpRemainder(i) = PolyNumerator(i)
    Next i
    
    InvUpperDenominator = 1# / PolyDenominator(UpperBoundDenominator)
    k = UpperBoundDenominator - 1
'i is the index for Quotient
    For i = UpperBoundQuotient To 0 Step -1
        ratio = tmpRemainder(i + UpperBoundDenominator) * InvUpperDenominator
        PolyQuotient(i) = ratio
        tmpRemainder(i + UpperBoundDenominator) = 0#
'j is the index for Denominator
        For j = LowerBoundDenominator To k
            tmpRemainder(i + j) = tmpRemainder(i + j) - ratio * PolyDenominator(j)
        Next j
    Next i
    For i = 0 To UpperBoundRemainder
        PolyRemainder(i) = tmpRemainder(i)
    Next i
    For i = UpperBoundQuotient + 1 To UBound(PolyQuotient)
        PolyQuotient(i) = 0#
    Next i
    For i = UpperBoundRemainder + 1 To UBound(PolyRemainder)
        PolyRemainder(i) = 0#
    Next i
End If

End Sub
' Range of Atn2: 0 <= Atn2 < 2 * pi
'
Public Function Atn2(x As Double, y As Double) As Double
Dim pi As Double

pi = 4# * Atn(1#)

'Prevent "Overflow" and "Devided by zero"
If Abs(y) > Abs(1000000# * x) Or x = 0# Then
    If y > 0# Then
        Atn2 = 0.5 * pi
    ElseIf y < 0# Then
        Atn2 = 1.5 * pi
    Else
        Atn2 = 0#
    End If
    Exit Function
End If

If x < 0# Then
'x < 0
    Atn2 = Atn(y / x) + pi
Else
    If y < 0# Then
'x > 0 and y < 0
        Atn2 = Atn(y / x) + 2 * pi
    Else
'x > 0 and y >= 0
        Atn2 = Atn(y / x)
    End If
End If

End Function
