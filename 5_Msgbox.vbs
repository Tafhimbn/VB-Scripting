' msgbox in VBScript
' ============================================================================
' A message box is a dialog box that displays a message to the user.
' The MsgBox function is used to create a message box in VBScript.
' The syntax for the MsgBox function is as follows:
' MsgBox(String[prompt], "buttons", "title", [helpfile], [context])
' Parameters:
    ' prompt: The message displayed in the message box. This parameter is required.
    ' buttons: A numeric expression that specifies the buttons and icons to be displayed in the message box. This parameter is optional.
        'In VBScript, MsgBox returns a number depending on which button the user clicks.
        'Common Return Values
            ' 1 → OK (vbOKOnly, vbOKCancel)
            ' 2 → Cancel (vbOKCancel, vbAbortRetryIgnore, vbRetryCancel)
            ' 3 → Abort (vbAbortRetryIgnore)
            ' 4 → Retry (vbAbortRetryIgnore, vbRetryCancel)
            ' 5 → Ignore (vbAbortRetryIgnore)
            ' 6 → Yes (vbYesNo, vbYesNoCancel)
            ' 7 → No (vbYesNo, vbYesNoCancel)
            ' so on...
    ' title: The title of the message box window. This parameter is optional.
    ' helpfile: The name of the help file to be displayed. This parameter is optional.
    ' context: The context ID for the help file. This parameter is optional.
' The MsgBox function returns an integer value that indicates which button was clicked by the user.
' If the user closes the message box, the function returns 0.

Dim message, title, response1, response2, response3
message = "Hello, this is a message box!" ' The message to be displayed in the message box
title = "Message Box Title" ' The title of the message box window   
'Basic Example:
response1 = MsgBox(message) ' Display the message box with OK button and Information icon
WScript.Echo "You clicked button number: " & response1 ' Print the button number clicked by the user
msgbox "You clicked button number: " & response1

response2 = MsgBox(message, 1) ' Display the message box with OK button and Cancel button
WScript.Echo "You clicked button number: " & response2 ' Print the button number clicked by the user
msgbox "You clicked button number: " & response2    

response3 = MsgBox(message, 2, title) ' Display the message box with OK button, Information icon and custom title
WScript.Echo "You clicked button number: " & response3 ' Print the button number clicked by the user
msgbox "You clicked button number: " & response3


