' Arrays in VBScript
' ============================================================================
' Arrays are used to store multiple values in a single variable.
' They can be single-dimensional or multi-dimensional.

' 1. Declaring and Initializing Arrays
' Syntax:
'        Dim arrayName(size)
'        arrayName = Array(value1, value2, ..., valueN)
'
' Example:
Dim fruits(2) ' Declare an array with 3 elements (0 to 2)   
fruits(0) = "Apple"
fruits(1) = "Banana"
fruits(2) = "Cherry"
WScript.Echo "Declared Array:"
For i = 0 To UBound(fruits)
    WScript.Echo fruits(i)
Next

' Using Array function to initialize an array
Dim colors
colors = Array("Red", "Green", "Blue")
WScript.Echo "Initialized Array using Array function:"
For i = 0 To UBound(colors) 
    WScript.Echo colors(i)                  
Next
WScript.Echo "---------------------"

' 2. Accessing Array Elements
' Syntax:
'        arrayName(index)
'        UBound(arrayName) → returns the upper bound (last index)
'        LBound(arrayName) → returns the lower bound (first index, usually 0)
WScript.Echo "Accessing Array Elements:"
WScript.Echo "Fruits Array:"
For i = LBound(fruits) To UBound(fruits)
    WScript.Echo "Index " & i & ": " & fruits(i)
Next
WScript.Echo "Colors Array:"
For i = LBound(colors) To UBound(colors)
    WScript.Echo "Index " & i & ": " & colors(i)
Next
WScript.Echo "---------------------"
' 3. Resizing Arrays
' Syntax:
'        ReDim arrayName(newSize)
'        ReDim Preserve arrayName(newSize) → preserves existing values
WScript.Echo "Resizing Arrays:"
ReDim fruits(4) ' Resize to hold 5 elements (0 to 4)
fruits(3) = "Date"
fruits(4) = "Elderberry"
WScript.Echo "Resized Fruits Array:"
For i = LBound(fruits) To UBound(fruits)
    WScript.Echo "Index " & i & ": " & fruits(i)
Next
WScript.Echo "---------------------"
