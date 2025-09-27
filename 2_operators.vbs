' Operators in VBScript 

'type of operators: 

'i) Arithmetic:
Dim a, b
a = 10  
b = 20
Dim sum, diff, product, quotient, modulus, exponent

'a) Addition "+":
sum = a + b  
WScript.Echo sum
msgbox "Addition: "&sum
'b) Subtraction "-":
diff = b - a 
WScript.Echo diff
msgbox "Subtraction: " & diff
'c) Multiplication "*":
product = a * b 
WScript.Echo product
msgbox "Multiplication: " & product
'd) Division "/":
quotient = b / a 
WScript.Echo quotient
msgbox "Division: " & quotient
'e) Modulus "Mod" or "%":
modulus = b Mod a
WScript.Echo modulus
msgbox "Modulus: " & modulus
'f) Exponentiation "^":
exponent = a ^ 2
WScript.Echo exponent
msgbox "Exponentiation: " & exponent

'ii) Comparison
  ' a) Equal to "="
    If a = b Then
        WScript.Echo "a is equal to b"
        msgbox "a is equal to b"
    Else
        WScript.Echo "a is not equal to b"
        msgbox "a is not equal to b"
    End If
  ' b) Not equal to "<>"
    If a <> b Then
        WScript.Echo "a is not equal to b"
        msgbox "a is not equal to b"
    Else
        WScript.Echo "a is equal to b"
        msgbox "a is equal to b"
    End If
    ' c) Greater than ">"
    If a > b Then
        WScript.Echo "a is greater than b"
        msgbox "a is greater than b"
    Else
        WScript.Echo "a is not greater than b"
        msgbox "a is not greater than b"
    End If
    ' d) Less than "<"
    If a < b Then
        WScript.Echo "a is less than b"
        msgbox "a is less than b"
    Else
        WScript.Echo "a is not less than b"
        msgbox "a is not less than b"
    End If
    ' e) Greater than or equal to ">="
    If a >= b Then
        WScript.Echo "a is greater than or equal to b"
        msgbox "a is greater than or equal to b"
    Else
        WScript.Echo "a is not greater than or equal to b"
        msgbox "a is not greater than or equal to b"
    End If
    ' f) Less than or equal to "<="
    If a <= b Then
        WScript.Echo "a is less than or equal to b"
        msgbox "a is less than or equal to b"
    Else
        WScript.Echo "a is not less than or equal to b"
        msgbox "a is not less than or equal to b"
    End If

'iii) Concatenation
Dim firstName, lastName, fullName
firstName = "Tafhim"
lastName = "Bin Nasir"
fullName = firstName & " " & lastName
WScript.Echo fullName
msgbox "Full Name: " & fullName

'iv) Logical
'a) And
If a < b And b = 20 Then
    WScript.Echo "And: Both conditions are true"
    msgbox "And: Both conditions are true"
Else
    WScript.Echo "And: One or both conditions are false"
    msgbox "And: One or both conditions are false"
End If  
'b) Or
If a < b Or b = 10 Then
    WScript.Echo "Or: At least one condition is true"
    msgbox "Or: At least one condition is true"
Else
    WScript.Echo "Or: Both conditions are false"
    msgbox "Or: Both conditions are false"
End If          
'c) Not
If Not (a > b) Then
    WScript.Echo "Not: Condition is true"
    msgbox "Not: Condition is true"
Else
    WScript.Echo "Not: Condition is false"
    msgbox "Not: Condition is false"
End If
'd) Xor
If a < b Xor b = 10 Then
    WScript.Echo "Xor: Only one condition is true"
    msgbox "Xor: Only one condition is true"
Else
    WScript.Echo "Xor: Both conditions are either true or false"
    msgbox "Xor: Both conditions are either true or false"
End If
