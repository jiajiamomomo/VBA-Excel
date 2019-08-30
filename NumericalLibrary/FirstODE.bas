Attribute VB_Name = "FirstODE"
Option Explicit
'solve first oder differential equation which is of the
'name of dydx_ for a uniform distributed grid.
Public Sub TestFirstODE()
Dim x0 As Double, y0 As Double
Dim xUpperBound As Double, dx As Double
Dim y() As Double, yy() As Double, yyy() As Double

x0 = 0#
y0 = 0#
xUpperBound = 10#
dx = 0.01
ReDim y(0 To (xUpperBound - x0) / dx)
ReDim yy(0 To (xUpperBound - x0) / dx)
ReDim yyy(0 To (xUpperBound - x0) / dx)
Call RungeKutta(x0, y0, xUpperBound, dx, y())
Call RungeKuttaMerson(x0, y0, xUpperBound, dx, yy())
Call ButcherRungeKutta(x0, y0, xUpperBound, dx, yyy())

Dim sh As Worksheet
Set sh = ThisWorkbook.Worksheets("FirstODE")
sh.Columns(1).Clear
sh.Columns(2).Clear
sh.Columns(3).Clear
Dim i As Long
For i = LBound(y) To UBound(y)
    sh.Cells(i - LBound(y) + 1, 1).Value = y(i)
    sh.Cells(i - LBound(y) + 1, 2).Value = yy(i)
    sh.Cells(i - LBound(y) + 1, 3).Value = yyy(i)
Next i

End Sub
'first order differential equation dy/dx = function(x, y)
'
Public Function dydx_(x As Double, y As Double) As Double
dydx_ = Exp(x) + y
End Function
'solve first order differential equation dy/dx = sub(x, y)
'error of order dx^4
Public Sub RungeKutta(x0 As Double, y0 As Double, _
                      xUpperBound As Double, dx As Double, _
                      y() As Double)
Dim x As Double, yy As Double
Dim k1 As Double, k2 As Double, k3 As Double, k4 As Double
Const OneSixth  As Double = 1# / 6#
Dim Half_dx As Double
Dim i As Long, l As Long

If ((xUpperBound - x0) / dx + 1) > (UBound(y) - LBound(y) + 1) Then
    MsgBox "The number of elements of parameter y() is not enough!"
End If
Half_dx = 0.5 * dx

x = x0
yy = y0
l = LBound(y)
i = l
y(i) = yy
For i = l + 1 To UBound(y)
    k1 = dx * dydx_(x, yy)
    k2 = dx * dydx_(x + Half_dx, yy + 0.5 * k1)
    k3 = dx * dydx_(x + Half_dx, yy + 0.5 * k2)
    k4 = dx * dydx_(x + dx, yy + k3)
    yy = yy + (k1 + 2# * k2 + 2# * k3 + k4) * OneSixth
    x = x + dx
    y(i) = yy
Next i

End Sub
'solve first order differential equation dy/dx = sub(x, y)
'error of order dx^5
Public Sub RungeKuttaMerson(x0 As Double, y0 As Double, _
                            xUpperBound As Double, dx As Double, _
                            y() As Double)
Dim x As Double, yy As Double
Dim k1 As Double, k2 As Double, k3 As Double
Dim k4 As Double, k5 As Double
Const OneThird  As Double = 1# / 3#
Const OneSixth  As Double = 1# / 6#
Dim OneThird_dx As Double
Dim Half_dx As Double
Dim i As Long, l As Long

OneThird_dx = dx / 3#
Half_dx = 0.5 * dx

x = x0
yy = y0
l = LBound(y)
i = l
y(i) = yy
For i = l + 1 To UBound(y)
    k1 = dx * dydx_(x, yy)
    k2 = dx * dydx_(x + OneThird_dx, yy + OneThird * k1)
    k3 = dx * dydx_(x + OneThird_dx, yy + OneSixth * k1 + OneSixth * k2)
    k4 = dx * dydx_(x + Half_dx, yy + 0.125 * k1 + 0.375 * k3)
    k5 = dx * dydx_(x + dx, yy + 0.5 * k1 - 1.5 * k3 + 2# * k4)
    yy = yy + (k1 + 4# * k4 + k5) * OneSixth
    x = x + dx
    y(i) = yy
Next i

End Sub
'solve first order differential equation dy/dx = sub(x, y)
'error of order dx^6
Public Sub ButcherRungeKutta(x0 As Double, y0 As Double, _
                             xUpperBound As Double, dx As Double, _
                             y() As Double)
Dim x As Double, yy As Double
Dim k1 As Double, k2 As Double, k3 As Double
Dim k4 As Double, k5 As Double, k6 As Double
Dim OneQuarter_dx As Double
Dim Half_dx As Double
Dim ThreeQuarter_dx As Double
'c1 = 0.077777777...
Const c1 As Double = 7# / 90#
'c2 = 0.0
'Const c2 As Double = 0#
'c3 = 0.355555555...
Const c3 As Double = 32# / 90#
'c4 = 0.133333333...
Const c4 As Double = 12# / 90#
'c5 = 0.355555555...
Const c5 As Double = 32# / 90#
'c6 = 0.077777777...
Const c6 As Double = 7# / 90#
'cc1 = -0.42857142857...
Const cc1 As Double = -3# / 7#
'cc2 = 0.2857142857
Const cc2 As Double = 2# / 7#
'cc3 = 1.7142857...
Const cc3 As Double = 12# / 7#
'cc4 = -1.7142857...
Const cc4 As Double = -12# / 7#
'cc5 = 1.142857142857...
Const cc5 As Double = 8# / 7#
Dim i As Long, l As Long

OneQuarter_dx = 0.25 * dx
Half_dx = 0.5 * dx
ThreeQuarter_dx = 0.75 * dx

x = x0
yy = y0
l = LBound(y)
i = l
y(i) = yy
For i = l + 1 To UBound(y)
    k1 = dx * dydx_(x, yy)
    k2 = dx * dydx_(x + OneQuarter_dx, yy + 0.25 * k1)
    k3 = dx * dydx_(x + OneQuarter_dx, yy + 0.125 * k1 + 0.125 * k2)
    k4 = dx * dydx_(x + Half_dx, yy - 0.5 * k2 + k3)
    k5 = dx * dydx_(x + ThreeQuarter_dx, yy + 0.1875 * k1 + 0.5625 * k4)
    k6 = dx * dydx_(x + dx, yy + cc1 * k1 _
                               + cc2 * k2 _
                               + cc3 * k3 _
                               + cc4 * k4 _
                               + cc5 * k5)
    yy = yy + (c1 * k1 + c3 * k3 + _
               c4 * k4 + c5 * k5 + c6 * k6)
    x = x + dx
    y(i) = yy
Next i

End Sub
