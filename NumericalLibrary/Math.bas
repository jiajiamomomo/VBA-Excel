Attribute VB_Name = "Math"
Option Explicit
'可推導數學函數
'
'Secant(正割)
'unit of X is radian
Public Function Sec(x As Double) As Double
Sec = 1# / Cos(x)
End Function
'Cosecant(餘割)
'unit of X is radian
Public Function Csc(x As Double) As Double
Csc = 1# / Sin(x)
End Function
'Cotangent(餘切)
'unit of X is radian
Public Function Cot(x As Double) As Double
Cot = 1# / Tan(x)
End Function
'Inverse Sine(反正弦)
'unit of Asin is radian
Public Function Asin(x As Double) As Double
If Abs(x) > 1# Then
    MsgBox "The value of parameter X should be between -1 and 1."
    Exit Function
End If
Asin = Atn(x / Sqr(-x * x + 1#))
End Function
'Inverse Cosine(反餘弦)
'unit of Acos is radian
Public Function Acos(x As Double) As Double
If Abs(x) > 1# Then
    MsgBox "The value of parameter X should be between -1 and 1."
    Exit Function
End If
Acos = Atn(-x / Sqr(-x * x + 1#)) + 2# * Atn(1#)
End Function
'Inverse Secant(反正割)
'unit of Asec is radian
Public Function Asec(x As Double) As Double
If Abs(x) < 1# Then
    MsgBox "The value of parameter X should be larger than or equal to 1 or less than or equal to -1."
    Exit Function
End If
Asec = Atn(x / Sqr(x * x - 1#)) + Sgn((x) - 1#) * (2# * Atn(1#))
End Function
'Inverse Cosecant(反餘割)
'unit of Acsc is radian
Public Function Acsc(x As Double) As Double
If Abs(x) < 1# Then
    MsgBox "The value of parameter X should be larger than or equal to 1 or less than or equal to -1."
    Exit Function
End If
Acsc = Atn(x / Sqr(x * x - 1#)) + (Sgn(x) - 1#) * (2# * Atn(1#))
End Function
'Inverse Cotangent(反餘切)
'unit of Acot is radian
Public Function Acot(x As Double) As Double
Acot = Atn(x) + 2# * Atn(1#)
End Function
'Hyperbolic Sine(超正弦)
Public Function HSin(x As Double) As Double
Dim ExpX As Double
ExpX = Exp(x)
HSin = 0.5 * (ExpX - 1# / ExpX)
End Function
'Hyperbolic Cosine(超餘弦)
Public Function HCos(x As Double) As Double
Dim ExpX As Double
ExpX = Exp(x)
HCos = 0.5 * (ExpX + 1# / ExpX)
End Function
'Hyperbolic Tangent(超正切)
Public Function HTan(x As Double) As Double
Dim ExpX As Double, ExpMX As Double
ExpX = Exp(x)
ExpMX = 1# / Exp(x)
HTan = (ExpX - ExpMX) / (ExpX + ExpMX)
End Function
'Hyperbolic Secant(超正割)
Public Function HSec(x As Double) As Double
Dim ExpX As Double
ExpX = Exp(x)
HSec = 2# / (ExpX + 1# / ExpX)
End Function
'Hyperbolic Cosecant(超餘割)
Public Function Hcsc(x As Double) As Double
Dim ExpX As Double
ExpX = Exp(x)
Hcsc = 2# / (ExpX - 1# / ExpX)
End Function
'Hyperbolic Cotangent(超餘切)
Public Function HCot(x As Double) As Double
Dim ExpX As Double, ExpMX As Double
ExpX = Exp(x)
ExpMX = 1# / Exp(x)
HCot = (ExpX + ExpMX) / (ExpX - ExpMX)
End Function
'Inverse Hyperbolic Sine(反超正弦)
Public Function Hasin(x As Double) As Double
Hasin = Log(x + Sqr(x * x + 1#))
End Function
'Inverse Hyperbolic Cosine(反超餘弦)
Public Function Hacos(x As Double) As Double
If x < 1# Then
    MsgBox "The value of parameter X should be larger than or equal to 1."
    Exit Function
End If
Hacos = Log(x + Sqr(x * x - 1#))
End Function
'Inverse Hyperbolic Tangent(反超正切)
Public Function Hatan(x As Double) As Double
If Abs(x) >= 1# Then
    MsgBox "The value of parameter X should be between -1 and 1."
    Exit Function
End If
Hatan = 0.5 * Log((1# + x) / (1# - x))
End Function
'Inverse Hyperbolic Secant(反超正割)
Public Function Hasec(x As Double) As Double
If x > 1# Or x <= 0# Then
    MsgBox "The value of parameter X should be between 0 and 1."
    Exit Function
End If
Hasec = Log((Sqr(-x * x + 1#) + 1#) / x)
End Function
'Inverse Hyperbolic Cosecant(反超餘割)
Public Function Hacsc(x As Double) As Double
If x = 0# Then
    MsgBox "The value of parameter X should not be 0."
    Exit Function
End If
Hacsc = Log((Sgn(x) * Sqr(x * x + 1#) + 1#) / x)
End Function
'Inverse Hyperbolic Cotangent(反超餘切)
Public Function Hacot(x As Double) As Double
If Abs(x) <= 1# Then
    MsgBox "The value of parameter X should be larger than or equal to 1 or less than or equal to -1."
    Exit Function
End If
Hacot = 0.5 * Log((x + 1#) / (x - 1#))
End Function
'以 N 為底的對數值
Public Function LogN(x As Double, n As Double) As Double
If x <= 0# Or n <= 0# Then
    MsgBox "The value of parameter X and parameter n should be larger than 0."
    Exit Function
End If
LogN = Log(x) / Log(n)
End Function

