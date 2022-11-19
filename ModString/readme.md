
# ModString
Description: String operation.<br>
License: MIT, Free Software<br>

## Srting Operation in specified Range/Selection
*Parameter:*<br>
*rngRange: Range containing String value*<br>

Sub RngStr2UpperCase(rngRange)<br>

Sub SelStr2UpperCase()<br>

Sub RngStr2LowerCase(rngRange)<br>

Sub SelStr2LowerCase()<br>

Sub RngStr2SentenceCase(rngRange)<br>

Sub SelStr2SentenceCase()<br>

Sub RngStr2CapitalCase(rngRange)<br>

Sub SelStr2CapitalCase()<br>

Sub RngStr2Trim(rngRange)<br>

Sub SelStr2Trim()<br>

Sub RngStr2LTrim(rngRange)<br>

Sub SelStr2LTrim()<br>

Sub RngStr2RTrim(rngRange)<br>

Sub SelStr2RTrim()<br>


## String Operation for String variable
*Parameter:*<br>
*text_string: string to be processed*<br>
*sString: string to be processed*<br>

Function RemoveChar(text_string, char_removed)<br>

Function RemoveASC(text_string, Ascii_removed)<br>

*index starts from 0*<br>
Function SplitString(text_string, delimiter, index)<br>

*Check if every character in sString is from 'A' to 'Z'*<br>
*if sString is empty string, return False*<br>
Function IsUpper(sString) As Boolean<br>

*Check if every character in sString is from 'a' to 'z'*<br>
*if sString is empty string, return False*<br>
Function IsLower(sString) As Boolean<br>

*Check if every character in sString is from 'A' to 'Z' or from 'a' to 'z'*<br>
*if sString is empty string, return False*<br>
Function IsLetter(sString) As Boolean<br>

*Check if every character in sString is from '0' to '9'*<br>
*if sString is empty string, return False*<br>
Function IsDigit(sString) As Boolean<br>

Function RepeatString(sString, number)<br>
