Attribute VB_Name = "ModStringDev"
'
'String Functions & Subs
'

'Like operator depends on Option Compare
'Option Compare Text     'Case-insensitive
Option Compare Binary   'Case-sensitive, default

'Using VBScript.RegExp
'In VBA IDE, select "Tools" -> "Reference" -> "Microsoft VBScript Regular Expressions 5.5"



'slower
'about 0.14s
'repeat IsUpper(Ucase(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10))) for 500,000 times
Private Function IsUpper_For(sString) As Boolean
'Check if every character in sString is from 'A' to 'Z'
iLen = Len(sString)
If iLen = 0 Then
    IsUpper_For = False
    Exit Function
End If

IsUpper_For = True
For i = 1 To iLen
    c = Mid(sString, i, 1)
    If c < "A" Then
        IsUpper_For = False
        Exit Function
    ElseIf c > "Z" Then
        IsUpper_For = False
        Exit Function
    End If
Next
End Function

'slower
'about 0.14s
'repeat IsLower(Lcase(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10))) for 500,000 times
Private Function IsLower_For(sString) As Boolean
'Check if every character in sString is from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsLower_For = False
    Exit Function
End If

IsLower_For = True
For i = 1 To iLen
    c = Mid(sString, i, 1)
    If c < "a" Then
        IsLower_For = False
        Exit Function
    ElseIf c > "z" Then
        IsLower_For = False
        Exit Function
    End If
Next
End Function

'slower
'about 36s
'repeat IsLetter(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10)) for 500,000 times
Private Function IsLetter_For(sString) As Boolean
'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsLetter_For = False
    Exit Function
End If

IsLetter_For = True
For intPos = 1 To iLen
    Select Case Asc(Mid(sString, intPos, 1))
        Case 65 To 90, 97 To 122    ''A' to 'Z' or from 'a' to 'z'
            'do nothing
        Case Else
            IsLetter_For = False
            Exit For
    End Select
Next
End Function

'slower
'about 24s
'repeat IsDigit(RepeatString("963852741", 50)) for 500,000 times
Private Function IsDigit_Select(sString) As Boolean
'Check if every character in sString is from '0' to '9'
iLen = Len(sString)
If iLen = 0 Then
    IsDigit_Select = False
    Exit Function
End If

IsDigit_Select = True
For intPos = 1 To iLen
    Select Case Asc(Mid(sString, intPos, 1))
        Case 48 To 57    ''0' to '9'
            'do nothing
        Case Else
            IsDigit_Select = False
            Exit For
    End Select
Next
End Function

'slower
'about 31s
'repeat IsDigit(RepeatString("963852741", 50)) for 500,000 times
Private Function IsDigit_If1(sString) As Boolean
'Check if every character in sString is from '0' to '9'
iLen = Len(sString)
If iLen = 0 Then
    IsDigit_If1 = False
    Exit Function
End If

IsDigit_If1 = True
For i = 1 To iLen
    c = Mid(sString, i, 1)
    'check incorrect case
    If c < "0" Or c > "9" Then
        IsDigit_If1 = False
        Exit Function
    End If
Next
End Function

'slower
'about 34s
'repeat IsDigit(RepeatString("963852741", 50)) for 500,000 times
Private Function IsDigit_If2(sString) As Boolean
'Check if every character in sString is from '0' to '9'
iLen = Len(sString)
If iLen = 0 Then
    IsDigit_If2 = False
    Exit Function
End If

IsDigit_If2 = True
For i = 1 To iLen
    c = Mid(sString, i, 1)
    'check correct case
    If c >= "0" And c <= "9" Then
        'correct
    Else
        'incorrect
        IsDigit_If2 = False
        Exit Function
    End If
Next
End Function



'fastest
'about 0.07s
'repeat IsUpper(Ucase(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10))) for 500,000 times
Private Function IsUpper_Like(sString) As Boolean
'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsUpper_Like = False
    Exit Function
End If

IsUpper_Like = sString Like RepeatString("[A-Z]", iLen)
End Function

'fastest
'about 0.07s
'repeat IsLower(Lcase(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10))) for 500,000 times
Private Function IsLower_Like(sString) As Boolean
'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsLower_Like = False
    Exit Function
End If

IsLower_Like = sString Like RepeatString("[a-z]", iLen)
End Function

'fastest
'about 24s
'repeat IsLetter(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10)) for 500,000 times
Private Function IsLetter_Like(sString) As Boolean
'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsLetter_Like = False
    Exit Function
End If

IsLetter_Like = sString Like RepeatString("[a-zA-Z]", iLen)
End Function

'fastest
'about 0.7s
'repeat IsDigit(RepeatString("963852741", 50)) for 500,000 times
Private Function IsDigit_Like(sString) As Boolean
'Check if every character in sString is from '0' to '9'
iLen = Len(sString)
If iLen = 0 Then
    IsDigit_Like = False
    Exit Function
End If

IsDigit_Like = sString Like String(iLen, "#")
End Function



'Using VBScript.RegExp
'In VBA IDE, select "Tools" -> "Reference" -> "Microsoft VBScript Regular Expressions 5.5"

'slowest
'about 370s
'repeat IsUpper(Ucase(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10))) for 500,000 times
Private Function IsUpper_RegExp(sString) As Boolean
Dim RegEx
Dim iLen
Set RegEx = CreateObject("VBScript.RegExp")
        
'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsUpper_RegExp = False
    Exit Function
End If

RegEx.IgnoreCase = False
RegEx.Pattern = "[A-Z]{" & iLen & "}"
IsUpper_RegExp = RegEx.Test(sString)
End Function

'slowest
'about 370s
'repeat IsLower(Lcase(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10))) for 500,000 times
Private Function IsLower_RegExp(sString) As Boolean
Dim RegEx
Dim iLen
Set RegEx = CreateObject("VBScript.RegExp")
        
'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsLower_RegExp = False
    Exit Function
End If

RegEx.IgnoreCase = False
RegEx.Pattern = "[a-z]{" & iLen & "}"
IsLower_RegExp = RegEx.Test(sString)
End Function

'slowest
'about 380s
'repeat IsLetter(RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10)) for 500,000 times
Private Function IsLetter_RegExp(sString) As Boolean
Dim RegEx
Dim iLen
Set RegEx = CreateObject("VBScript.RegExp")
        
'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
iLen = Len(sString)
If iLen = 0 Then
    IsLetter_RegExp = False
    Exit Function
End If

RegEx.IgnoreCase = True
RegEx.Pattern = "[A-Z]{" & iLen & "}"
IsLetter_RegExp = RegEx.Test(sString)
End Function

'slowest
'about 380s
'repeat IsDigit(RepeatString("963852741", 50)) for 500,000 times
Private Function IsDigit_RegExp(sString) As Boolean
Dim RegEx
Dim iLen
Set RegEx = CreateObject("VBScript.RegExp")

'Check if every character in sString is from '0' to '9'
iLen = Len(sString)
If iLen = 0 Then
    IsDigit_RegExp = False
    Exit Function
End If

RegEx.Pattern = "\d{" & iLen & "}"
IsDigit_RegExp = RegEx.Test(sString)
End Function



'copied from
'https://codereview.stackexchange.com/questions/159080/string-repeat-function-in-vba
'fast
'about 4.2s
'repeat RepeatString("a-z", 100) for 500,000 times
Private Function RepeatString_Mid(sString, number)
Dim s As String
Dim c As Long
Dim l As Long
Dim i As Long

l = Len(sString)
c = l * number
s = Space$(c)

For i = 1 To c Step l
    Mid(s, i, l) = sString
Next

RepeatString_Mid = s
End Function

'slow
'about 5.8s
'repeat RepeatString("a-z", 100) for 500,000 times
Private Function RepeatString_Concatenate(sString, number)
Dim s As String
s = ""

For i = 1 To number
    s = s & sString
Next

RepeatString_Concatenate = s
End Function

Private Sub TestRepeatString()
Dim i, j As Integer
Dim iNum, jNum As Integer
Dim s
Dim TimerStart, TimerEnd
Dim TimerLoop, TimerRepeatString_Mid, TimerRepeatString_Concatenate

s = RepeatString_Mid("", 100)
MsgBox s
MsgBox Len(s)
s = RepeatString_Concatenate("", 100)
MsgBox s
MsgBox Len(s)

iNum = 1000
jNum = 500

'Application.ScreenUpdating = False
'Application.EnableEvents = False

'dummy loop body
s = "dummy loop body"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        
    Next
Next
TimerEnd = Timer
TimerLoop = TimerEnd - TimerStart

'RepeatString_Mid
s = "RepeatString_Mid"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        s = RepeatString_Mid("a-z", 100)
    Next
Next
TimerEnd = Timer
TimerRepeatString_Mid = (TimerEnd - TimerStart) - TimerLoop

'RepeatString_Concatenate
s = "RepeatString_Concatenate"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        s = RepeatString_Concatenate("a-z", 100)
    Next
Next
TimerEnd = Timer
TimerRepeatString_Concatenate = (TimerEnd - TimerStart) - TimerLoop

'Application.ScreenUpdating = True
'Application.EnableEvents = True

Application.StatusBar = ""
MsgBox "Repeat for " & (iNum * jNum) & " times" & Chr(13) & _
       "Dummy loop body takes " & TimerLoop & " seconds" & Chr(13) & _
       "RepeatString_Mid takes " & TimerRepeatString_Mid & " seconds" & Chr(13) & _
       "RepeatString_Concatenate takes " & TimerRepeatString_Concatenate & " seconds"
End Sub

Private Sub testStringAsIs()
Dim i, j As Integer
Dim iNum, jNum As Integer
Dim s, sString
Dim b
Dim TimerStart, TimerEnd
Dim TimerLoop, TimerString_For, TimerString_If1, TimerString_If2, TimerString_Select
Dim TimerString_Regex, TimerString_Like

sString = RepeatString("963852741", 50)
MsgBox "IsDigit function should all be True" & Chr(13) & _
       IsDigit_If1(sString) & Chr(13) & _
       IsDigit_If2(sString) & Chr(13) & _
       IsDigit_Select(sString) & Chr(13) & _
       IsDigit_Like(sString) & Chr(13) & _
       IsDigit_RegExp(sString)

sString = RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10)
MsgBox "IsLetter function should all be True" & Chr(13) & _
       IsLetter_For(sString) & Chr(13) & _
       IsLetter_Like(sString) & Chr(13) & _
       IsLetter_RegExp(sString)

sString = UCase(sString)
MsgBox "IsUpper function should all be True" & Chr(13) & _
       IsUpper_For(sString) & Chr(13) & _
       IsUpper_Like(sString) & Chr(13) & _
       IsUpper_RegExp(sString)

sString = LCase(sString)
MsgBox "IsLower function should all be True" & Chr(13) & _
       IsLower_For(sString) & Chr(13) & _
       IsLower_Like(sString) & Chr(13) & _
       IsLower_RegExp(sString)

'Application.ScreenUpdating = False
'Application.EnableEvents = False

FileNumber = FreeFile
Open ThisWorkbook.Path & "\ModuleString.txt" For Output As FileNumber

iNum = 100
jNum = 500

'dummy loop body
s = "dummy loop body"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        
    Next
Next
TimerEnd = Timer
TimerLoop = TimerEnd - TimerStart



sString = RepeatString("963852741", 50)

'IsDigit_If1
s = "IsDigit_If1"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsDigit_If1(sString)
    Next
Next
TimerEnd = Timer
TimerString_If1 = (TimerEnd - TimerStart) - TimerLoop

'IsDigit_If2
s = "IsDigit_If2"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsDigit_If2(sString)
    Next
Next
TimerEnd = Timer
TimerString_If2 = (TimerEnd - TimerStart) - TimerLoop

'IsDigit_Select
s = "IsDigit_Select"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsDigit_Select(sString)
    Next
Next
TimerEnd = Timer
TimerString_Select = (TimerEnd - TimerStart) - TimerLoop

'IsDigit_Like
s = "IsDigit_Like"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsDigit_Like(sString)
    Next
Next
TimerEnd = Timer
TimerString_Like = (TimerEnd - TimerStart) - TimerLoop

'IsDigit_RegExp
s = "IsDigit_RegExp"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsDigit_RegExp(sString)
    Next
Next
TimerEnd = Timer
TimerString_Regex = (TimerEnd - TimerStart) - TimerLoop

Application.StatusBar = ""
mesg = "Repeat for " & (iNum * jNum) & " times" & vbNewLine & _
       "Dummy loop body takes " & TimerLoop & " seconds" & vbNewLine & _
       "IsDigit_If1 takes " & TimerString_If1 & " seconds" & vbNewLine & _
       "IsDigit_If2 takes " & TimerString_If2 & " seconds" & vbNewLine & _
       "IsDigit_Select takes " & TimerString_Select & " seconds" & vbNewLine & _
       "IsDigit_Like takes " & TimerString_Like & " seconds" & vbNewLine & _
       "IsDigit_RegExp takes " & TimerString_Regex & " seconds"
Print #FileNumber,
Print #FileNumber, mesg



sString = RepeatString("aAbBcCdDeEfFgGhHiIjJkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ", 10)

'IsLetter_For
s = "IsLetter_For"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsLetter_For(sString)
    Next
Next
TimerEnd = Timer
TimerString_For = (TimerEnd - TimerStart) - TimerLoop

'IsLetter_Like
s = "IsLetter_Like"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsLetter_Like(sString)
    Next
Next
TimerEnd = Timer
TimerString_Like = (TimerEnd - TimerStart) - TimerLoop

'IsLetter_RegExp
s = "IsLetter_RegExp"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsLetter_RegExp(sString)
    Next
Next
TimerEnd = Timer
TimerString_Regex = (TimerEnd - TimerStart) - TimerLoop

Application.StatusBar = ""
mesg = "Repeat for " & (iNum * jNum) & " times" & vbNewLine & _
       "Dummy loop body takes " & TimerLoop & " seconds" & vbNewLine & _
       "IsLetter_For takes " & TimerString_For & " seconds" & vbNewLine & _
       "IsLetter_Like takes " & TimerString_Like & " seconds" & vbNewLine & _
       "IsLetter_RegExp takes " & TimerString_Regex & " seconds"
Print #FileNumber,
Print #FileNumber, mesg



sString = UCase(sString)

'IsUpper_For
s = "IsUpper_For"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsUpper_For(sString)
    Next
Next
TimerEnd = Timer
TimerString_For = (TimerEnd - TimerStart) - TimerLoop

'IsUpper_Like
s = "IsUpper_Like"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsUpper_Like(sString)
    Next
Next
TimerEnd = Timer
TimerString_Like = (TimerEnd - TimerStart) - TimerLoop

'IsUpper_RegExp
s = "IsUpper_RegExp"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsUpper_RegExp(sString)
    Next
Next
TimerEnd = Timer
TimerString_Regex = (TimerEnd - TimerStart) - TimerLoop

Application.StatusBar = ""
mesg = "Repeat for " & (iNum * jNum) & " times" & vbNewLine & _
       "Dummy loop body takes " & TimerLoop & " seconds" & vbNewLine & _
       "IsUpper_For takes " & TimerString_For & " seconds" & vbNewLine & _
       "IsUpper_Like takes " & TimerString_Like & " seconds" & vbNewLine & _
       "IsUpper_RegExp takes " & TimerString_Regex & " seconds"
Print #FileNumber,
Print #FileNumber, mesg



sString = LCase(sString)

'IsLower_For
s = "IsLower_For"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsLower_For(sString)
    Next
Next
TimerEnd = Timer
TimerString_For = (TimerEnd - TimerStart) - TimerLoop

'IsLower_Like
s = "IsLower_Like"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsLower_Like(sString)
    Next
Next
TimerEnd = Timer
TimerString_Like = (TimerEnd - TimerStart) - TimerLoop

'IsLower_RegExp
s = "IsLower_RegExp"
Application.StatusBar = s
Debug.Print s
TimerStart = Timer
For i = 1 To iNum
    For j = 1 To jNum
        b = IsLower_RegExp(sString)
    Next
Next
TimerEnd = Timer
TimerString_Regex = (TimerEnd - TimerStart) - TimerLoop

Application.StatusBar = ""
mesg = "Repeat for " & (iNum * jNum) & " times" & vbNewLine & _
       "Dummy loop body takes " & TimerLoop & " seconds" & vbNewLine & _
       "IsLower_For takes " & TimerString_For & " seconds" & vbNewLine & _
       "IsLower_Like takes " & TimerString_Like & " seconds" & vbNewLine & _
       "IsLower_RegExp takes " & TimerString_Regex & " seconds"
Print #FileNumber,
Print #FileNumber, mesg

Debug.Print "Done"

s = "Python"
mesg = "s = '" & s & "'" & vbNewLine & _
       "IsUpper(s) is " & IsUpper(s) & vbNewLine & _
       "IsLower(s) is " & IsLower(s) & vbNewLine & _
       "IsLetter(s) is " & IsLetter(s) & vbNewLine & _
       "IsDigit(s) is " & IsDigit(s)
Print #FileNumber,
Print #FileNumber, mesg

s = UCase("Python")
mesg = "s = '" & s & "'" & vbNewLine & _
       "IsUpper(s) is " & IsUpper(s) & vbNewLine & _
       "IsLower(s) is " & IsLower(s) & vbNewLine & _
       "IsLetter(s) is " & IsLetter(s) & vbNewLine & _
       "IsDigit(s) is " & IsDigit(s)
Print #FileNumber,
Print #FileNumber, mesg

s = LCase("Python")
mesg = "s = '" & s & "'" & vbNewLine & _
       "IsUpper(s) is " & IsUpper(s) & vbNewLine & _
       "IsLower(s) is " & IsLower(s) & vbNewLine & _
       "IsLetter(s) is " & IsLetter(s) & vbNewLine & _
       "IsDigit(s) is " & IsDigit(s)
Print #FileNumber,
Print #FileNumber, mesg

s = "963852741"
mesg = "s = '" & s & "'" & vbNewLine & _
       "IsUpper(s) is " & IsUpper(s) & vbNewLine & _
       "IsLower(s) is " & IsLower(s) & vbNewLine & _
       "IsLetter(s) is " & IsLetter(s) & vbNewLine & _
       "IsDigit(s) is " & IsDigit(s)
Print #FileNumber,
Print #FileNumber, mesg

Close #FileNumber

'Application.ScreenUpdating = True
'Application.EnableEvents = True

End Sub



Sub DebugClear()
Application.SendKeys "^g ^a {DEL}"

'Application.VBE.Windows("Immediate").SetFocus
'If Application.VBE.ActiveWindow.Caption = "Immediate" And _
'   Application.VBE.ActiveWindow.Visible Then
'    Application.SendKeys "^a {DEL} {HOME}"
'End If
End Sub



