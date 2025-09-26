' VBscript Stands for Visual Basic Scripting. It is a scripting language.
'It works on Windows OS.


' Dim is used declaration of variables.
' Case insensitive Eg: Dim VALUE|
' Variable names cannot contain spaces Eg: Dim My Value (Invalid)
' Variable names cannot start with a number Eg: Dim 1Value (Invalid)
' Variable names can include letters, numbers, and the underscore character Eg: Dim My_Value (Valid)
' Variable names cannot include special characters like !, @, #, $, %, ^, &, *, (, ), -, +, =, {, }, [, ], |, \, :, ;, ", ', <, >, ,, ., ?, / Eg: Dim My$Value (Invalid)
' VBScript does not support block or multi-line comments. Each line of a comment must start with a single quote (').
' VBScript is not case-sensitive, meaning that variable names and keywords are treated the same regardless of their case.

Dim value ' This declares a variable named value
value = 10 ' This assigns the value 10 to the variable value
WScript.Echo value ' This prints the value of the variable value to the console
msgbox "The value is " & value ' This shows a message box with the value of the variable value
' You can also use the MsgBox function to display a message box with a custom message.
' The & operator is used to concatenate (join) strings in VBScript.

nonDeclaredVariable = 20 ' This assigns the value 20 to a variable that has not been declared
WScript.Echo nonDeclaredVariable ' This prints the value of the non-declared variable to the console
msgbox "The non-declared variable is " & nonDeclaredVariable ' This shows a message box with the value of the non-declared variable
' In VBScript, you can assign values to variables without declaring them first. However, it is a good practice to declare variables using the Dim statement to avoid potential issues with variable scope and naming conflicts.
' VBScript does not have strict data types, so you can assign different types of values to the same variable without any type declaration.


