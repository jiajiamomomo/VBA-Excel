Attribute VB_Name = "Interpolation"
Option Explicit
'
'
Public Sub TestInterpolation()
Dim i As Integer, xLM1 As Integer, yLM1 As Integer
Dim l As Integer, m As Integer, n As Integer
Dim x0 As Double
Dim y0 As Double
Const xL As Double = 1.5
Const h As Double = 0.01
Dim x() As Double
Dim y() As Double
Dim UseEqualSpace As Boolean

UseEqualSpace = True
n = 100
xLM1 = -23
yLM1 = -43

ReDim x(1 + xLM1 To n + xLM1)
ReDim y(1 + yLM1 To n + yLM1)
Randomize timer
For i = 1 To n
    If UseEqualSpace Then
        x(i + xLM1) = xL + CDbl(i - 1) * h
    Else
        x(i + xLM1) = Rnd()
    End If
    y(i + yLM1) = Rnd()
Next i
Call QuickSort(x())

m = 5 + Int((n - 5 - 5 + 1) * Rnd())

ThisWorkbook.Worksheets("Interpolation").Activate
Cells(2, 2).Value = x(m - 1 + xLM1)
Cells(2, 3).Value = y(m - 1 + yLM1)
Cells(3, 2).Value = x(m + xLM1)
Cells(3, 3).Value = y(m + yLM1)
Cells(4, 2).Value = x(m + 1 + xLM1)
Cells(4, 3).Value = y(m + 1 + yLM1)
Cells(5, 2).Value = x(m + 2 + xLM1)
Cells(5, 3).Value = y(m + 2 + yLM1)

Columns(5).Clear
Columns(6).Clear
Columns(7).Clear
Columns(8).Clear
Cells(1, 5).Value = "x"
Cells(1, 6).Value = "y1"
Cells(1, 7).Value = "y2"
Cells(1, 8).Value = "y3"
For i = 0 To 10
    x0 = x(m + xLM1) + CDbl(i) / 10# * (x(m + 1 + xLM1) - x(m + xLM1))
    Cells(2 + i, 5).Value = x0
    
    If UseEqualSpace Then
        l = EqualSpaceLinearInterpolation(x(), y(), x0, y0)
    Else
        l = LinearInterpolation(x(), y(), x0, y0)
    End If
    Cells(2 + i, 6).Value = y0
    
    If UseEqualSpace Then
        l = EqualSpaceQuadraticInterpolation(x(), y(), x0, y0)
    Else
        l = QuadraticInterpolation(x(), y(), x0, y0)
    End If
    Cells(2 + i, 7).Value = y0
    
    If UseEqualSpace Then
        l = EqualSpaceCubicInterpolation(x(), y(), x0, y0)
    Else
        l = CubicInterpolation(x(), y(), x0, y0)
    End If
    Cells(2 + i, 8).Value = y0
Next i

End Sub
'Parameter array x() is assumed to be ordered.
'x(i) <= x(i+1)
Public Function LinearInterpolation(x() As Double, _
                                    y() As Double, _
                                    xInterpolate As Double, _
                                    yInterpolate As Double) As Integer
Dim i As Integer, iL As Integer, iU As Integer
Dim iP1 As Integer
Dim UpperBound As Integer, UpperBoundM1 As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim Xn As Double, XnP1 As Double
Dim xMXn As Double, xMXnP1 As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Function
End If
UpperBoundM1 = UpperBound - 1

'extrapolation
If xInterpolate < x(1 + xLM1) Then
    i = 1
    iP1 = 2
    Xn = x(i + xLM1)
    XnP1 = x(iP1 + xLM1)
    xMXn = xInterpolate - Xn
    xMXnP1 = xInterpolate - XnP1
    LinearInterpolation = xLM1
'extrapolation
ElseIf xInterpolate > x(UpperBound + xLM1) Then
    i = UpperBoundM1
    iP1 = UpperBound
    Xn = x(i + xLM1)
    XnP1 = x(iP1 + xLM1)
    xMXn = xInterpolate - Xn
    xMXnP1 = xInterpolate - XnP1
    LinearInterpolation = 1 + UpperBound + xLM1
'interpolation
Else
    iL = 1
    iU = UpperBound
    Do
        i = Int((iL + iU) / 2)
        iP1 = i + 1
        Xn = x(i + xLM1)
        XnP1 = x(iP1 + xLM1)
        xMXn = xInterpolate - Xn
        xMXnP1 = xInterpolate - XnP1
        If xMXn * xMXnP1 <= 0# Then
            LinearInterpolation = i + xLM1
            Exit Do
        End If
        If xMXn < 0# Then
            iU = i
        Else
            iL = iP1
        End If
    Loop
End If

yInterpolate = ((xMXnP1) * y(i + yLM1) - _
                (xMXn) * y(iP1 + yLM1)) / _
               (Xn - XnP1)

End Function
'Parameter array x() is assumed to be equal spaced.
'x(i+1) - x(i) = h
Public Function EqualSpaceLinearInterpolation(x() As Double, _
                                              y() As Double, _
                                              xInterpolate As Double, _
                                              yInterpolate As Double) As Integer
Dim i As Integer, iP1 As Integer
Dim UpperBound As Integer, UpperBoundM1 As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim h As Double
Dim Xn As Double, XnP1 As Double
Dim xMXn As Double, xMXnP1 As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Function
End If
UpperBoundM1 = UpperBound - 1

Xn = x(1 + xLM1)
'extrapolation
If xInterpolate < Xn Then
    i = 1
    iP1 = 2
    xMXn = xInterpolate - x(i + xLM1)
    EqualSpaceLinearInterpolation = xLM1
'extrapolation
ElseIf xInterpolate > x(UpperBound + xLM1) Then
    i = UpperBoundM1
    iP1 = UpperBound
    xMXn = xInterpolate - x(i + xLM1)
    EqualSpaceLinearInterpolation = 1 + UpperBound + xLM1
'interpolation
Else
    h = x(2 + xLM1) - Xn
    i = 1 + Int((xInterpolate - Xn) / h)
    iP1 = i + 1
    Xn = x(i + xLM1)
    XnP1 = x(iP1 + xLM1)
    xMXn = xInterpolate - Xn
    xMXnP1 = xInterpolate - XnP1
    If xInterpolate < Xn Then
        i = i - 1
        iP1 = i + 1
        xMXn = xMXn + h
    ElseIf xInterpolate > XnP1 Then
        i = iP1
        iP1 = i + 1
        xMXn = xMXn - h
    End If
    EqualSpaceLinearInterpolation = i + xLM1
End If

yInterpolate = ((xMXn) * y(iP1 + yLM1) - _
                (xMXn - h) * y(i + yLM1)) / h

End Function
'Parameter array x() is assumed to be ordered.
'x(i) <= x(i+1)
Public Function QuadraticInterpolation(x() As Double, _
                                       y() As Double, _
                                       xInterpolate As Double, _
                                       yInterpolate As Double) As Integer
Dim i As Integer, iL As Integer, iU As Integer
Dim iP1 As Integer
Dim UpperBound As Integer, UpperBoundM1 As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim Xn As Double, XnP1 As Double
Dim Yn As Double, YnP1 As Double
Dim xMXn As Double, xMXnP1 As Double
Dim a As Double, b As Double, c As Double
Dim y0 As Double, y1 As Double, y2 As Double
Dim x1 As Double, x2 As Double, xin As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Function
End If
UpperBoundM1 = UpperBound - 1

If xInterpolate <= x(2 + xLM1) Then
    If xInterpolate < x(1 + xLM1) Then
        QuadraticInterpolation = xLM1
    Else
        QuadraticInterpolation = 1 + xLM1
    End If
    y0 = x(1 + xLM1)
    y1 = x(2 + xLM1)
    y2 = x(3 + xLM1)
    x1 = y1 - y0
    x2 = y2 - y0
    xin = xInterpolate - y0
    y0 = y(1 + yLM1)
    y1 = y(2 + yLM1)
    y2 = y(3 + yLM1)
    y1 = y1 - y0
    y2 = y2 - y0
    c = 1# / ((x1 - x2) * x1 * x2)
    a = (y1 * x2 - x1 * y2) * c
    b = (x1 * x1 * y2 - y1 * x2 * x2) * c
    c = y0
    yInterpolate = a * xin * xin + b * xin + c
    Exit Function
End If

If xInterpolate >= x(UpperBoundM1 + xLM1) Then
    If xInterpolate > x(UpperBound + xLM1) Then
        QuadraticInterpolation = 1 + UpperBound + xLM1
    Else
        QuadraticInterpolation = UpperBound + xLM1
    End If
    y0 = x(UpperBound - 2 + xLM1)
    y1 = x(UpperBoundM1 + xLM1)
    y2 = x(UpperBound + xLM1)
    x1 = y1 - y0
    x2 = y2 - y0
    xin = xInterpolate - y0
    y0 = y(UpperBound - 2 + yLM1)
    y1 = y(UpperBoundM1 + yLM1)
    y2 = y(UpperBound + yLM1)
    y1 = y1 - y0
    y2 = y2 - y0
    c = 1# / ((x1 - x2) * x1 * x2)
    a = (y1 * x2 - x1 * y2) * c
    b = (x1 * x1 * y2 - y1 * x2 * x2) * c
    c = y0
    yInterpolate = a * xin * xin + b * xin + c
    Exit Function
End If

iL = 2
iU = UpperBoundM1
Do
    i = Int((iL + iU) / 2)
    iP1 = i + 1
    Xn = x(i + xLM1)
    XnP1 = x(iP1 + xLM1)
    xMXn = xInterpolate - Xn
    xMXnP1 = xInterpolate - XnP1
    If xMXn * xMXnP1 <= 0# Then
        QuadraticInterpolation = i + xLM1
        If xInterpolate = Xn Then
            yInterpolate = y(i + yLM1)
            Exit Function
        ElseIf xInterpolate = XnP1 Then
            yInterpolate = y(iP1 + yLM1)
            Exit Function
        Else
            Yn = y(i + yLM1)
            YnP1 = y(iP1 + yLM1)
            
            y0 = x(i - 1 + xLM1)
            y1 = Xn
            y2 = XnP1
            x1 = y1 - y0
            x2 = y2 - y0
            xin = xInterpolate - y0
            y0 = y(i - 1 + yLM1)
            y1 = Yn
            y2 = YnP1
            y1 = y1 - y0
            y2 = y2 - y0
            c = 1# / ((x1 - x2) * x1 * x2)
            a = (y1 * x2 - x1 * y2) * c
            b = (x1 * x1 * y2 - y1 * x2 * x2) * c
            c = y0
            yInterpolate = a * xin * xin + b * xin + c
            
            y0 = Xn
            y1 = XnP1
            y2 = x(i + 2 + xLM1)
            x1 = y1 - y0
            x2 = y2 - y0
            xin = xMXn
            y0 = Yn
            y1 = YnP1
            y2 = y(i + 2 + yLM1)
            y1 = y1 - y0
            y2 = y2 - y0
            c = 1# / ((x1 - x2) * x1 * x2)
            a = (y1 * x2 - x1 * y2) * c
            b = (x1 * x1 * y2 - y1 * x2 * x2) * c
            c = y0
            yInterpolate = 0.5 * (yInterpolate + _
                           a * xin * xin + b * xin + c)
            Exit Function
        End If
    End If
    If xMXn < 0# Then
        iU = i
    Else
        iL = iP1
    End If
Loop

End Function
'Parameter array x() is assumed to be equal spaced.
'x(i+1) - x(i) = h
Public Function EqualSpaceQuadraticInterpolation(x() As Double, _
                                                 y() As Double, _
                                                 xInterpolate As Double, _
                                                 yInterpolate As Double) As Integer
Dim i As Integer, iP1 As Integer
Dim UpperBound As Integer, UpperBoundM1 As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim h As Double
Dim Xn As Double, XnP1 As Double
Dim Yn As Double, YnP1 As Double
Dim xMXn As Double, xMXnP1 As Double
Dim a As Double, b As Double, c As Double
Dim y0 As Double, y1 As Double, y2 As Double
Dim xin As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Function
End If
UpperBoundM1 = UpperBound - 1

Xn = x(1 + xLM1)
If xInterpolate <= x(2 + xLM1) Then
    If xInterpolate <= x(2 + xLM1) Then
        EqualSpaceQuadraticInterpolation = xLM1
    Else
        EqualSpaceQuadraticInterpolation = 1 + xLM1
    End If
    xin = xInterpolate - Xn
    y0 = y(1 + yLM1)
    y1 = y(2 + yLM1)
    y2 = y(3 + yLM1)
    h = 1# / h
    a = 0.5 * (y2 - 2# * y1 + y0) * h * h
    b = 0.5 * (4# * y1 - y2 - 3# * y0) * h
    c = y0
    yInterpolate = a * xin * xin + b * xin + c
    Exit Function
End If

If xInterpolate >= x(UpperBoundM1 + xLM1) Then
    If xInterpolate > x(UpperBound + xLM1) Then
        EqualSpaceQuadraticInterpolation = 1 + UpperBound + xLM1
    Else
        EqualSpaceQuadraticInterpolation = UpperBound + xLM1
    End If
    xin = xInterpolate - x(UpperBoundM1 + xLM1) + h
    y0 = y(UpperBound - 2 + yLM1)
    y1 = y(UpperBoundM1 + yLM1)
    y2 = y(UpperBound + yLM1)
    h = 1# / h
    a = 0.5 * (y2 - 2# * y1 + y0) * h * h
    b = 0.5 * (4# * y1 - y2 - 3# * y0) * h
    c = y0
    yInterpolate = a * xin * xin + b * xin + c
    Exit Function
End If

h = x(2 + xLM1) - Xn
i = 1 + Int((xInterpolate - Xn) / h)
iP1 = i + 1
Xn = x(i + xLM1)
XnP1 = x(iP1 + xLM1)
xMXn = xInterpolate - Xn
xMXnP1 = xInterpolate - XnP1
If xMXn * xMXnP1 <= 0# Then

ElseIf xInterpolate < Xn Then
    i = i - 1
    iP1 = i + 1
    xMXn = xMXn + h
Else
    i = iP1
    iP1 = i + 1
    xMXn = xMXn - h
End If

EqualSpaceQuadraticInterpolation = i + xLM1
If xInterpolate = Xn Then
    yInterpolate = y(i + yLM1)
    Exit Function
ElseIf xInterpolate = XnP1 Then
    yInterpolate = y(iP1 + yLM1)
    Exit Function
Else
    Yn = y(i + yLM1)
    YnP1 = y(iP1 + yLM1)
    
    xin = xMXn + h
    y0 = y(i - 1 + yLM1)
    y1 = Yn
    y2 = YnP1
    h = 1# / h
    a = 0.5 * (y2 - 2# * y1 + y0) * h * h
    b = 0.5 * (4# * y1 - y2 - 3# * y0) * h
    c = y0
    yInterpolate = a * xin * xin + b * xin + c
    
    xin = xMXn
    y0 = Yn
    y1 = YnP1
    y2 = y(i + 2 + yLM1)
    a = 0.5 * (y2 - 2# * y1 + y0) * h * h
    b = 0.5 * (4# * y1 - y2 - 3# * y0) * h
    c = y0
    yInterpolate = 0.5 * (yInterpolate + _
                   a * xin * xin + b * xin + c)
    Exit Function
End If

End Function
'Parameter array x() is assumed to be ordered.
'x(i) <= x(i+1)
Public Function CubicInterpolation(x() As Double, _
                                   y() As Double, _
                                   xInterpolate As Double, _
                                   yInterpolate As Double) As Integer
Dim i As Integer, iL As Integer, iU As Integer
Dim iP1 As Integer
Dim UpperBound As Integer, UpperBoundM1 As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim h As Double
Dim Xn As Double, XnP1 As Double
Dim xMXn As Double, xMXnP1 As Double
Dim a As Double, b As Double, c As Double, d As Double
Dim tmp1 As Double, tmp2 As Double, tmp3 As Double
Dim y0 As Double, y1 As Double, y2 As Double, y3 As Double
Dim x1 As Double, x2 As Double, x3 As Double, xin As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Function
End If
UpperBoundM1 = UpperBound - 1

If xInterpolate <= x(2 + xLM1) Then
    If xInterpolate < x(1 + xLM1) Then
        CubicInterpolation = xLM1
    Else
        CubicInterpolation = 1 + xLM1
    End If
    y0 = x(1 + xLM1)
    y1 = x(2 + xLM1)
    y2 = x(3 + xLM1)
    y3 = x(4 + xLM1)
    x1 = y1 - y0
    x2 = y2 - y0
    x3 = y3 - y0
    xin = xInterpolate - y0
    y0 = y(1 + yLM1)
    y1 = y(2 + yLM1)
    y2 = y(3 + yLM1)
    y3 = y(4 + yLM1)
    y1 = y1 - y0
    y2 = y2 - y0
    y3 = y3 - y0
ElseIf xInterpolate >= x(UpperBoundM1 + xLM1) Then
    If xInterpolate > x(UpperBound + xLM1) Then
        CubicInterpolation = 1 + UpperBound + xLM1
    Else
        CubicInterpolation = UpperBound + xLM1
    End If
    y0 = x(UpperBound - 3 + xLM1)
    y1 = x(UpperBound - 2 + xLM1)
    y2 = x(UpperBoundM1 + xLM1)
    y3 = x(UpperBound + xLM1)
    x1 = y1 - y0
    x2 = y2 - y0
    x3 = y3 - y0
    xin = xInterpolate - y0
    y0 = y(UpperBound - 3 + yLM1)
    y1 = y(UpperBound - 2 + yLM1)
    y2 = y(UpperBoundM1 + yLM1)
    y3 = y(UpperBound + yLM1)
    y1 = y1 - y0
    y2 = y2 - y0
    y3 = y3 - y0
Else
    iL = 2
    iU = UpperBoundM1
    Do
        i = Int((iL + iU) / 2)
        iP1 = i + 1
        Xn = x(i + xLM1)
        XnP1 = x(iP1 + xLM1)
        xMXn = xInterpolate - Xn
        xMXnP1 = xInterpolate - XnP1
        If xMXn * xMXnP1 <= 0# Then
            CubicInterpolation = i + xLM1
            If xInterpolate = Xn Then
                yInterpolate = y(i + yLM1)
                Exit Function
            ElseIf xInterpolate = XnP1 Then
                yInterpolate = y(iP1 + yLM1)
                Exit Function
            Else
                y0 = x(i - 1 + xLM1)
                y1 = Xn
                y2 = XnP1
                y3 = x(i + 2 + xLM1)
                x1 = y1 - y0
                x2 = y2 - y0
                x3 = y3 - y0
                xin = xInterpolate - y0
                y0 = y(i - 1 + yLM1)
                y1 = y(i + yLM1)
                y2 = y(iP1 + yLM1)
                y3 = y(i + 2 + yLM1)
                y1 = y1 - y0
                y2 = y2 - y0
                y3 = y3 - y0
                Exit Do
            End If
        End If
        If xMXn < 0# Then
            iU = i
        Else
            iL = iP1
        End If
    Loop
End If

tmp1 = (x2 - x3) * x2 * x3
tmp2 = (x3 - x1) * x3 * x1
tmp3 = (x1 - x2) * x1 * x2
d = 1# / ((tmp1 + tmp2 + tmp3) * x1 * x2 * x3)
tmp1 = tmp1 * y1
tmp2 = tmp2 * y2
tmp3 = tmp3 * y3
a = (tmp1 + tmp2 + tmp3) * d
b = -((x3 + x2) * tmp1 + _
      (x1 + x3) * tmp2 + _
      (x2 + x1) * tmp3) * d
c = (x3 * x2 * tmp1 + _
     x1 * x3 * tmp2 + _
     x2 * x1 * tmp3) * d
d = y0
yInterpolate = a * xin * xin * xin + _
               b * xin * xin + _
               c * xin + _
               d

End Function
'Parameter array x() is assumed to be equal spaced.
'x(i+1) - x(i) = h
Public Function EqualSpaceCubicInterpolation(x() As Double, _
                                             y() As Double, _
                                             xInterpolate As Double, _
                                             yInterpolate As Double) As Integer
Dim i As Integer, iP1 As Integer
Dim UpperBound As Integer, UpperBoundM1 As Integer
Dim xLM1 As Integer, yLM1 As Integer
Dim h As Double
Dim Xn As Double, XnP1 As Double
Dim xMXn As Double, xMXnP1 As Double
Dim a As Double, b As Double, c As Double, d As Double
Dim y0 As Double, y1 As Double, y2 As Double, y3 As Double
Dim x1 As Double, x2 As Double, x3 As Double, xin As Double

xLM1 = LBound(x) - 1
yLM1 = LBound(y) - 1
UpperBound = UBound(x) - xLM1
If UpperBound <> (UBound(y) - yLM1) Then
    MsgBox "Number of element of Parameter array x() " & vbNewLine & _
           "is not equal to that of parameter array y()!"
    Exit Function
End If
UpperBoundM1 = UpperBound - 1

Xn = x(1 + xLM1)
If xInterpolate <= x(2 + xLM1) Then
    If xInterpolate < x(1 + xLM1) Then
        EqualSpaceCubicInterpolation = xLM1
    Else
        EqualSpaceCubicInterpolation = 1 + xLM1
    End If
    xin = xInterpolate - Xn
    y0 = y(1 + yLM1)
    y1 = y(2 + yLM1)
    y2 = y(3 + yLM1)
    y3 = y(4 + yLM1)
ElseIf xInterpolate >= x(UpperBoundM1 + xLM1) Then
    If xInterpolate > x(UpperBound + xLM1) Then
        EqualSpaceCubicInterpolation = 1 + UpperBound + xLM1
    Else
        EqualSpaceCubicInterpolation = UpperBound + xLM1
    End If
    xin = xInterpolate - x(UpperBound - 3 + xLM1)
    y0 = y(UpperBound - 3 + yLM1)
    y1 = y(UpperBound - 2 + yLM1)
    y2 = y(UpperBound - 1 + yLM1)
    y3 = y(UpperBound + yLM1)
Else
    h = x(2 + xLM1) - Xn
    i = 1 + Int((xInterpolate - Xn) / h)
    iP1 = i + 1
    Xn = x(i + xLM1)
    XnP1 = x(iP1 + xLM1)
    xMXn = xInterpolate - Xn
    xMXnP1 = xInterpolate - XnP1
    If xInterpolate < Xn Then
        i = i - 1
        iP1 = i + 1
        xMXn = xMXn + h
    ElseIf xInterpolate > XnP1 Then
        i = iP1
        iP1 = i + 1
        xMXn = xMXn - h
    End If
    
    EqualSpaceCubicInterpolation = i + xLM1
    If xInterpolate = Xn Then
        yInterpolate = y(i + yLM1)
        Exit Function
    ElseIf xInterpolate = XnP1 Then
        yInterpolate = y(iP1 + yLM1)
        Exit Function
    Else
        xin = xMXn + h
        y0 = y(i - 1 + yLM1)
        y1 = y(i + yLM1)
        y2 = y(iP1 + yLM1)
        y3 = y(i + 2 + yLM1)
    End If
End If

h = 1# / h
d = 1# / 6
a = (y3 - 3# * y2 + 3# * y1 - y0) * _
    d * h * h * h
b = (-3# * y3 + 12# * y2 - 15# * y1 + 6# * y0) * _
    d * h * h
c = (2# * y3 - 9# * y2 + 18# * y1 - 11# * y0) * _
    d * h
d = y0
yInterpolate = a * xin * xin * xin + _
               b * xin * xin + _
               c * xin + _
               d

End Function

