' Variable Declaration and Assignment in VBScript
' ============================================================================
' This script demonstrates how to declare and assign variables in VBScript.

Dim myVariable ' This declares a variable named myVariable
myVariable = 5 ' This assigns the value 5 to the variable myVariable
WScript.Echo myVariable ' This prints the value of the variable myVariable to the console
msgbox "The value of myVariable is " & myVariable ' This shows a message box with the value of the variable myVariable
' You can also use the MsgBox function to display a message box with a custom message.
' The & operator is used to concatenate (join) strings in VBScript. 

' rules for naming variables:
' 1. Variable names cannot contain spaces Eg: Dim My Value (Invalid)
' 2. Variable names cannot start with a number Eg: Dim 1Value (Invalid)
' 3. Variable names can include letters, numbers, and the underscore character Eg: Dim My_Value (Valid)
' 4. Variable names cannot include special characters like !, @, #, $, %, ^, &, *, (, ), -, +, =, {, }, [, ], |, \, :, ;, ", ', <, >, ,, ., ?, / Eg: Dim My$Value (Invalid)
' 5. VBScript does not support block or multi-line comments. Each line of a comment must start with a single quote (').
' 6. VBScript is not case-sensitive, meaning that variable names and keywords are treated the same regardless of their case.
' 7. It is a good practice to declare variables using the Dim statement to avoid potential issues with variable scope and naming conflicts.
' 8. VBScript does not have strict data types, so you can assign different types of values to the same variable without any type declaration.

'Types of variables in VBScript:

'Implicitly Typed Variables: These are variables that are declared without a specific data type. VBScript automatically assigns the Variant data type to these variables. Eg: Dim myVariable

'Explicitly Typed Variables: VBScript does not support explicitly typed variables like some other programming languages (e.g., Dim myVariable As Integer). All variables in VBScript are of the Variant data type by default.
' 1. String: Used to store text. Eg: Dim myString: myString = "Hello, World!"   
Dim myname = "Tafhim Bin NAsir"
WScript.Echo myname
msgbox "My name is " & myname

' 2. Integer: Used to store whole numbers. Eg: Dim myInteger: myInteger = 42
Dim mybirthYear = 1998
WScript.Echo mybirthYear
msgbox "My birth year is " & mybirthYear
' 3. Long: Used to store larger whole numbers. Eg: Dim myLong: myLong = 1234567890
Dim myNumber = 1234567890   
WScript.Echo myNumber
msgbox "My number is " & myNumber
' 4. Single: Used to store single-precision floating-point numbers. Eg: Dim mySingle: mySingle = 3.14
Dim age = 26.5
WScript.Echo age
msgbox "My age is " & age
' 5. Double: Used to store double-precision floating-point numbers. Eg: Dim myDouble: myDouble = 3.14159265358979
Dim pi = 3.14159265358979
WScript.Echo pi
msgbox "The value of pi is " & pi
' 6. Currency: Used to store monetary values. Eg: Dim myCurrency: myCurrency = 19.99
' 7. Date: Used to store date and time values. Eg: Dim myDate: myDate = #12/31/2023#
' 8. Boolean: Used to store True or False values. Eg: Dim myBoolean : myBoolean = True
' 9. Variant: The default data type that can hold any type of data. Eg  : Dim myVariant: myVariant = "Hello" or myVariant = 100 or myVariant = 3.14
' 10. Object: Used to store references to objects. Eg: Dim myObject: Set myObject = CreateObject("Scripting.FileSystemObject")
