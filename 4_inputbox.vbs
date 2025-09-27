' Input Box in VBScript
' ============================================================================
' An Input Box is a dialog box that prompts the user to enter information.
' The InputBox function is used to create an input box in VBScript.
' The syntax for the InputBox function is as follows:
' InputBox(String[prompt], String[title], String[default input], [long xpos], [long ypos], [String[helpfile], [long context]])
' Parameters:
    ' prompt: The message displayed in the input box. This parameter is required.
    ' title: The title of the input box window. This parameter is optional.
    ' default: The default text displayed in the input box. This parameter is optional.
    ' xpos: The horizontal position of the input box. This parameter is optional.
    ' ypos: The vertical position of the input box. This parameter is optional.
    ' helpfile: The name of the help file to be displayed. This parameter is optional.
    ' context: The context ID for the help file. This parameter is optional.
' The InputBox function returns the text entered by the user as a string.
' If the user clicks the Cancel button, the function returns an empty string ("").  
Dim userInput1, userInput2, userInput3, userInput4
userInput1 = InputBox("Please enter your first name:")   
WScript.Echo "You entered: " & userInput1
msgbox "You entered: " & userInput1

userInput2 = InputBox("Please enter your last name:", "User Input")   
WScript.Echo "You entered: " & userInput2
msgbox "You entered: " & userInput2

userInput3 = InputBox("Please enter your :", "User Input","Ex. 25")   
WScript.Echo "You entered: " & userInput3
msgbox "You entered: " & userInput3

userInput4 = InputBox("Please enter your :", "User Input","Ex. 25", 200, 200)  
WScript.Echo "You entered: " & userInput4
msgbox "You entered: " & userInput4
