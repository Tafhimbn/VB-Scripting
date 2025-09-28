' Loops in VBScript
' ============================================================================
'Types of loops:
' 1. For...Next → Used for iterating a specific number of times
' 2. For Each...Next → Used for iterating over a collection or array
' 3. Do...Loop → Used for looping until a condition is met
' 4. While...Wend → Used for looping while a condition is true
' ============================================================================

' 1. For...Next Loop
' Syntax:
'        For counter = start To end [Step stepValue]
'            ' Statements
'        Next
'
'        counter → loop variable
'        start / end → start and end values
'        Step → optional (default = 1; can be negative)

WScript.Echo "For...Next Loop:"
For i = 1 To 5
    WScript.Echo "Iteration " & i
Next    
WScript.Echo "---------------------"

' 2. For Each...Next Loop
' Syntax:
'        For Each element In collection
'            ' Statements
'        Next

'        element → loop variable
'        collection → array or collection to iterate over

WScript.Echo "For Each...Next Loop:"
Dim arr
arr = Array("Apple", "Banana", "Cherry")

For Each fruit In arr
    WScript.Echo fruit
Next
WScript.Echo "---------------------"

' 3. Do...Loop
' Syntax:
' a) Do While...Loop (check before)
'        Do [While condition | Until condition]
'            ' Statements
'        Loop
    WScript.Echo "Do...Loop:"
    count = 1

    Do While count <= 5
        WScript.Echo "Count is " & count
        count = count + 1
    Loop
    WScript.Echo "---------------------"

' b) Do...Loop While (check after)
'        Do
'            ' Statements
'        Loop [While condition | Until condition]

'        condition → loop continues while true (While) or until true (Until)
    num = 1
    Do
        WScript.Echo "Number is " & num
        num = num + 1
    Loop While num <= 5 
    WScript.Echo "---------------------"

' c) Do Until...Loop (check before)
'        Do Until condition
'            ' Statements
'        Loop
'        condition → loop continues until true
    num = 1
    Do Until num > 5
        WScript.Echo "Number is " & num
        num = num + 1
    Loop
    WScript.Echo "---------------------"

' d) Do...Loop Until (check after)
'        Do
'            ' Statements
'        Loop Until condition
'        condition → loop continues until true
    num = 1
    Do 
        WScript.Echo "Number is " & num
        num = num + 1
    Loop Until num > 5
    WScript.Echo "---------------------"

' 4. While...Wend
WScript.Echo "While...Wend Loop:"
num = 1
While num <= 5
    WScript.Echo "Number is " & num
    num = num + 1
Wend
WScript.Echo "---------------------"
WScript.Echo "End of Loop Examples"
