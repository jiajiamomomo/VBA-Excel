Attribute VB_Name = "Integration"
Option Explicit
'The data in the x direction must be equally spaced for
'all methods of integration in this module!
Public Sub TestIntegration()
Dim NumInterval As Integer
Const xLM1 As Integer = -37
Const yLM1 As Integer = -39
Dim i As Integer, r As Integer
Dim poly(0 To 4) As Double
Dim polyIntegration(0 To 5) As Double
Dim x() As Double
Dim y() As Double
Const xStart As Double = 0#
Const xEnd As Double = 8#
Dim h As Double
Dim TrueIntegrate As Double
Dim TrapezoidalIntegrate As Double
Dim SimpsonIntegrate As Double
Dim Integrate As Double

'polynomial Test
'poly(0) = -4.16666666666667E-02
'poly(1) = 0.666666666666667
'poly(2) = -3.33333333333333
'poly(3) = 5.33333333333333
'poly(4) = 1
'Call PolynomialIntegrate(poly(), polyIntegration())
'TrueIntegrate = PolynomialValue(polyIntegration(), xEnd) _
'              - PolynomialValue(polyIntegration(), xStart)
'cos Test
TrueIntegrate = Sin(xEnd) - Sin(xStart)

ThisWorkbook.Worksheets("Integration").Activate
For r = 2 To 6
    NumInterval = Cells(r, 1).Value
    h = (xEnd - xStart) / NumInterval
    ReDim x(1 + xLM1 To NumInterval + 1 + xLM1)
    ReDim y(1 + yLM1 To NumInterval + 1 + yLM1)
                  
    For i = 1 To NumInterval + 1
        x(i + xLM1) = xStart + (i - 1) * h
'polynomial Test
'        y(i + yLM1) = PolynomialValue(poly(), x(i + xLM1))
'cos Test
        y(i + yLM1) = Cos(x(i + xLM1))
    Next i
    
'actual order of error is n ^ -2
    TrapezoidalIntegrate = TrapezoidalIntegration(h, y())
    Cells(r, 2).Value = Abs(TrapezoidalIntegrate - TrueIntegrate)
'actual order of error is n ^ -4
'NumInterval must be even then the number of data will
'be odd for this method!
    SimpsonIntegrate = SimpsonIntegration(h, y())
    Cells(r, 3).Value = Abs(SimpsonIntegrate - TrueIntegrate)
'actual order of error is n ^ -4
'NumInterval must be even then the number of data will
'be odd for this method!
    Integrate = Integration(h, y())
    Cells(r, 4).Value = Abs(Integrate - TrueIntegrate)
Next r

End Sub
'Trapezoidal rule
'Order of h ^ 3
Public Function TrapezoidalIntegration(h As Double, y() As Double) As Double
Dim yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer

yLM1 = LBound(y) - 1
UpperBound = UBound(y) - yLM1
If UpperBound < 2 Then
    MsgBox "Number of element of Parameter array y() " & vbNewLine & _
           "must be larger than or equal to 2!"
    Exit Function
End If

TrapezoidalIntegration = 0.5 * (y(1 + yLM1) + y(UpperBound + yLM1))
For i = 2 To UpperBound - 1
    TrapezoidalIntegration = TrapezoidalIntegration + y(i + yLM1)
Next i
TrapezoidalIntegration = TrapezoidalIntegration * h

End Function
'Simpson's rule
'Order of h ^ 5
'The number of data must be odd!
Public Function SimpsonIntegration(h As Double, y() As Double) As Double
Dim yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer
Dim y1 As Double, y3 As Double

yLM1 = LBound(y) - 1
UpperBound = UBound(y) - yLM1
If UpperBound Mod 2 <> 1 Then
    MsgBox "Number of element of Parameter array y() " & vbNewLine & _
           "must be odd!"
    Exit Function
End If
If UpperBound < 3 Then
    MsgBox "Number of element of Parameter array y() " & vbNewLine & _
           "must be larger than or equal to 3!"
    Exit Function
End If

y1 = y(1 + yLM1)
For i = 3 To UpperBound Step 2
    y3 = y(i + yLM1)
    SimpsonIntegration = SimpsonIntegration + y1 + _
                                              4# * y(i - 1 + yLM1) + _
                                              y3
    y1 = y3
Next i
SimpsonIntegration = SimpsonIntegration * h / 3#
End Function
'combine Simpson's 3/8 rule and Simpson's rule
'Order of n ^ -4
'The number of data must be odd and be larger than or equal to 9!
Private Function Integration(h As Double, y() As Double) As Double
Dim yLM1 As Integer
Dim UpperBound As Integer
Dim i As Integer

yLM1 = LBound(y) - 1
UpperBound = UBound(y) - yLM1
If UpperBound Mod 2 <> 1 Then
    MsgBox "Number of element of Parameter array y() " & vbNewLine & _
           "must be odd!"
    Exit Function
End If
If UpperBound < 9 Then
    MsgBox "Number of element of Parameter array y() " & vbNewLine & _
           "must be larger than or equal to 9!"
    Exit Function
End If

Integration = (17# * (y(1 + yLM1) + y(UpperBound + yLM1)) _
             + 59# * (y(2 + yLM1) + y(UpperBound - 1 + yLM1)) _
             + 43# * (y(3 + yLM1) + y(UpperBound - 2 + yLM1)) _
             + 49# * (y(4 + yLM1) + y(UpperBound - 3 + yLM1))) _
             / 48#
For i = 5 To UpperBound - 4
    Integration = Integration + y(i + yLM1)
Next i
Integration = Integration * h
End Function
