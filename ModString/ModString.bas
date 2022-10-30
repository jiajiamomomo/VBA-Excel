Attribute VB_Name = "ModString"
Option Explicit



Sub RngStr2UpperCase(rngRange)
If TypeName(rngRange) <> "Range" Then
    MsgBox "RngStr2UpperCase err: rngRange is not Range. type=" & TypeName(rngRange)
    Exit Sub
End If

For Each cell In rngRange
    If VarType(cell.Value) = vbString Then
        cell.Value = UCase(cell.Value)
    End If
Next
End Sub

Sub SelStr2UpperCase()
If TypeName(rngRange) = "Range" Then
Call RngStr2UpperCase(Selection)
End If
End Sub

Sub RngStr2LowerCase(rngRange)
If TypeName(rngRange) <> "Range" Then
    MsgBox "RngStr2LowerCase err: rngRange is not Range. type=" & TypeName(rngRange)
    Exit Sub
End If

For Each cell In rngRange
    If VarType(cell.Value) = vbString Then
        cell.Value = LCase(cell.Value)
    End If
Next
End Sub

Sub SelStr2LowerCase()
If TypeName(rngRange) = "Range" Then
Call RngStr2LowerCase(Selection)
End If
End Sub

Sub RngStr2SentenceCase(rngRange)
If TypeName(rngRange) <> "Range" Then
    MsgBox "RngStr2SentenceCase err: rngRange is not Range. type=" & TypeName(rngRange)
    Exit Sub
End If

For Each cell In rngRange
    If VarType(cell.Value) = vbString Then
        Dim first
        first = UCase(Left(cell.Value, 1))
        If Len(cell.Value) >= 2 Then
            cell.Value = first & Mid(cell.Value, 2)
        Else
            cell.Value = first
        End If
    End If
Next
End Sub

Sub SelStr2SentenceCase()
If TypeName(rngRange) = "Range" Then
Call RngStr2SentenceCase(Selection)
End If
End Sub

Sub RngStr2CapitalCase(rngRange)
If TypeName(rngRange) <> "Range" Then
    MsgBox "RngStr2CapitalCase err: rngRange is not Range. type=" & TypeName(rngRange)
    Exit Sub
End If

For Each cell In rngRange
    If VarType(cell.Value) = vbString Then
        Dim s
        s = vbNullString
        
        Dim CurrAsLower
        Dim CurrAsUpper
        Dim PrevAsLower
        Dim PrevAsUpper
        Dim counter
        For counter = 1 To Len(cell.Value)
            Dim char
            char = Mid(cell.Value, counter, 1)
            CurrAsLower = char Like "[a-z]"
            CurrAsUpper = char Like "[A-Z]"
            If counter = 1 Then
                If CurrAsLower Then
                    char = UCase(char)
                End If
            Else
                If (Not PrevAsLower) And (Not PrevAsUpper) And (CurrAsLower) Then
                    char = UCase(char)
                End If
            End If
            s = s & char
            PrevAsLower = CurrAsLower
            PrevAsUpper = CurrAsUpper
        Next
        cell.Value = s
    End If
Next
End Sub

Sub SelStr2CapitalCase()
If TypeName(rngRange) = "Range" Then
Call RngStr2CapitalCase(Selection)
End If
End Sub

Sub RngStr2Trim(rngRange)
If TypeName(rngRange) <> "Range" Then
    MsgBox "RngStr2Trim err: rngRange is not Range. type=" & TypeName(rngRange)
    Exit Sub
End If

For Each cell In rngRange
    If VarType(cell.Value) = vbString Then
        cell.Value = Trim(cell.Value)
    End If
Next
End Sub

Sub SelStr2Trim()
If TypeName(rngRange) = "Range" Then
Call RngStr2Trim(Selection)
End If
End Sub

Sub RngStr2LTrim(rngRange)
If TypeName(rngRange) <> "Range" Then
    MsgBox "RngStr2LTrim err: rngRange is not Range. type=" & TypeName(rngRange)
    Exit Sub
End If

For Each cell In rngRange
    If VarType(cell.Value) = vbString Then
        cell.Value = LTrim(cell.Value)
    End If
Next
End Sub

Sub SelStr2LTrim()
If TypeName(rngRange) = "Range" Then
Call RngStr2LTrim(Selection)
End If
End Sub

Sub RngStr2RTrim(rngRange)
If TypeName(rngRange) <> "Range" Then
    MsgBox "RngStr2RTrim err: rngRange is not Range. type=" & TypeName(rngRange)
    Exit Sub
End If

For Each cell In rngRange
    If VarType(cell.Value) = vbString Then
        cell.Value = RTrim(cell.Value)
    End If
Next
End Sub

Sub SelStr2RTrim()
If TypeName(rngRange) = "Range" Then
Call RngStr2RTrim(Selection)
End If
End Sub



'Function RangeStringValue(rngRange As Range) as String
Function RangeStringValue(rngRange)
If TypeName(rngRange) <> "Range" Then
    RangeStringValue = "RangeStringValue err: rngRange is not Range." & Chr(13) & _
                       "rngRange is " & TypeName(rngRange)
    Exit Function
End If


RangeStringValue = ""
For Each cell In rngRange
    RangeStringValue = RangeStringValue & cell.Value & Chr(13)
Next
End Function

'Function RangeStringText(rngRange As Range) as String
Function RangeStringText(rngRange)
If TypeName(rngRange) <> "Range" Then
    RangeStringText = "RangeStringText err: rngRange is not Range." & Chr(13) & _
                       "rngRange is " & TypeName(rngRange)
    Exit Function
End If


RangeStringText = ""
For Each cell In rngRange
    RangeStringText = RangeStringText & cell.Text & Chr(13)
Next
End Function

'Sub CharInString(text_string as String)
Sub CharInString(text_string)
If VarType(text_string) <> vbString Then
    MsgBox "CharInString err: text_string is not String." & Chr(13) & _
           "text_string is " & TypeName(text_string)
    Exit Sub
End If


msg = text_string & Chr(13)
For counter = 1 To Len(text_string)
    char = Mid(text_string, counter, 1)
    msg = msg & char & " " & Asc(char) & Chr(13)
Next
MsgBox msg
End Sub

'Sub ASCInString(text_string as String)
Sub ASCInString(text_string)
If VarType(text_string) <> vbString Then
    MsgBox "ASCInString err: text_string is not String." & Chr(13) & _
           "text_string is " & TypeName(text_string)
    Exit Sub
End If


msg = text_string & Chr(13)
For counter = 1 To Len(text_string)
    char = Mid(text_string, counter, 1)
    msg = msg & Asc(char) & Chr(13)
Next
MsgBox msg
End Sub

'Function RemoveChar(text_string As String, char_removed As String) As String
Function RemoveChar(text_string, char_removed)
returned_string = ""
If VarType(text_string) <> vbString Then
    returned_string = returned_string & "RemoveChar err: text_string is not String." & Chr(13) & _
                     "text_string is " & TypeName(text_string) & Chr(13)
End If
If VarType(char_removed) <> vbString Then
    returned_string = returned_string & "RemoveChar err: char_removed is not String." & Chr(13) & _
                      "char_removed is " & TypeName(char_removed) & Chr(13)
End If
If returned_string <> "" Then
    RemoveChar = returned_string
    Exit Function
End If


returned_string = ""
For i = 1 To Len(text_string)
    char_i = Mid(text_string, i, 1)
    If char_i <> char_removed Then
      returned_string = returned_string & char_i
    End If
Next
RemoveChar = returned_string
End Function

'Function RemoveASC(text_string As String, Ascii_removed As Integer) As String
Function RemoveASC(text_string, Ascii_removed)
ErrMsg = ""
If VarType(text_string) <> vbString Then
    ErrMsg = ErrMsg & "RemoveASC err: text_string is not String." & Chr(13) & _
             "text_string is " & TypeName(text_string) & Chr(13)
End If
If Not WorksheetFunction.IsNumber(Ascii_removed) Then
    ErrMsg = ErrMsg & "RemoveASC err: Ascii_removed is not number." & Chr(13) & _
             "Ascii_removed is " & TypeName(Ascii_removed) & Chr(13)
End If
If ErrMsg <> "" Then
    RemoveASC = ErrMsg
    Exit Function
End If

ErrMsg = ""
If (Ascii_removed < 0) Or (Ascii_removed > 255) Then
    ErrMsg = ErrMsg & "RemoveASC err: Ascii_removed is out of range." & Chr(13) & _
             "Ascii_removed=" & Ascii_removed & Chr(13)
End If
If ErrMsg <> "" Then
    RemoveASC = ErrMsg
    Exit Function
End If


char_removed = char(Ascii_removed)
RemoveASC = RemoveChar(text_string, char_removed)
End Function

'Function SplitString(text_string As String, delimiter As String, index As Integer) As String
Function SplitString(text_string, delimiter, index)
ErrMsg = ""
If VarType(text_string) <> vbString Then
    ErrMsg = ErrMsg & "SplitString err: text_string is not String." & Chr(13) & _
             "text_string is " & TypeName(text_string) & Chr(13)
End If
If VarType(delimiter) <> vbString Then
    ErrMsg = ErrMsg & "SplitString err: delimiter is not String." & Chr(13) & _
             "delimiter is " & TypeName(delimiter) & Chr(13)
End If
If Not WorksheetFunction.IsNumber(index) Then
    ErrMsg = ErrMsg & "SplitString err: index is not number." & Chr(13) & _
             "index is " & TypeName(index) & Chr(13)
End If
If ErrMsg <> "" Then
    SplitString = ErrMsg
    Exit Function
End If


Dim WrdArray() As String
WrdArray() = Split(text_string, delimiter)

ErrMsg = ""
If (index < LBound(WrdArray)) Or (index > UBound(WrdArray)) Then
    SplitString = "SplitString err: index is out of range." & Chr(13) & _
                  "index=" & index & Chr(13)
    Exit Function
End If


SplitString = WrdArray(index)   'index starts from 0
End Function



'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
'if sString is empty string, return False
'Function IsUpper(sString As String) As Boolean
Public Function IsUpper(sString) As Boolean
If VarType(sString) <> vbString Then
    IsUpper = "IsUpper err: sString is not String." & Chr(13) & _
                    "sString is " & TypeName(sString) & Chr(13)
    Exit Function
End If

Dim iLen
iLen = Len(sString)
If iLen = 0 Then
    IsUpper = False
    Exit Function
End If

IsUpper = sString Like RepeatString("[A-Z]", iLen)
End Function

'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
'if sString is empty string, return False
'Function IsLower(sString As String) As Boolean
Public Function IsLower(sString) As Boolean
If VarType(sString) <> vbString Then
    IsLower = "IsLower err: sString is not String." & Chr(13) & _
                    "sString is " & TypeName(sString) & Chr(13)
    Exit Function
End If

Dim iLen
iLen = Len(sString)
If iLen = 0 Then
    IsLower = False
    Exit Function
End If

IsLower = sString Like RepeatString("[a-z]", iLen)
End Function

'Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'
'if sString is empty string, return False
'Function IsLetter(sString As String) As Boolean
Public Function IsLetter(sString) As Boolean
If VarType(sString) <> vbString Then
    IsLetter = "IsLetter err: sString is not String." & Chr(13) & _
                         "sString is " & TypeName(sString) & Chr(13)
    Exit Function
End If

Dim iLen
iLen = Len(sString)
If iLen = 0 Then
    IsLetter = False
    Exit Function
End If

IsLetter = sString Like RepeatString("[a-zA-Z]", iLen)
End Function

'Check if every character in sString is from '0' to '9'
'if sString is empty string, return False
'Function IsDigit(sString As String) As Boolean
Public Function IsDigit(sString) As Boolean
If VarType(sString) <> vbString Then
    IsDigit = "IsDigit err: sString is not String." & Chr(13) & _
                      "sString is " & TypeName(sString) & Chr(13)
    Exit Function
End If

Dim iLen
iLen = Len(sString)
If iLen = 0 Then
    IsDigit = False
    Exit Function
End If

IsDigit = sString Like String(iLen, "#")
End Function



'copied from
'https://codereview.stackexchange.com/questions/159080/string-repeat-function-in-vba
'Function RepeatString(sString As String, number As Integer) As String
Public Function RepeatString(sString, number)
Dim ErrMsg
ErrMsg = ""
If VarType(sString) <> vbString Then
    ErrMsg = ErrMsg & "RepeatString err: sString is not String." & Chr(13) & _
             "sString is " & TypeName(sString) & Chr(13)
End If
If Not WorksheetFunction.IsNumber(number) Then
    ErrMsg = ErrMsg & "RepeatString err: number is not number." & Chr(13) & _
             "number is " & TypeName(number) & Chr(13)
End If
If ErrMsg <> "" Then
    RepeatString = ErrMsg
    Exit Function
End If

ErrMsg = ""
If number < 0 Then
    ErrMsg = ErrMsg & "RepeatString err: number is less than zero." & Chr(13) & _
             "number=" & number & Chr(13)
End If
If ErrMsg <> "" Then
    RepeatString = ErrMsg
    Exit Function
End If


If (sString = "") Or (number = 0) Then
    RepeatString = ""
    Exit Function
End If

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

RepeatString = s
End Function

