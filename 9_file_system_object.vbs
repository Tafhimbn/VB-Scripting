' File System Object in VBScript

' The File System Object (FSO) is a powerful tool in VBScript that allows you to
' interact with the file system on your computer. You can create, read, write,
' and delete files and folders, as well as retrieve information about them.

' ============================================================================

' 1. Create an instance of the File System Object
Set fso = CreateObject("Scripting.FileSystemObject")
    ' CreateObject("Scripting.FileSystemObject") → Creates a new instance of the File System Object.
    ' Set fso = ... → Stores that instance in the variable fso.
    ' At this point, you can use the fso variable to access various methods and properties of the File System Object.

' Example: Check if a file exists
If fso.FileExists("E:\CS, IOT, Embedded System\Github\VB Scripting\example.txt") Then
    WScript.Echo "File exists."
Else
    WScript.Echo "File does not exist."
End If

' Example: Create a new text file and write to it
Set file = fso.CreateTextFile("E:\CS, IOT, Embedded System\Github\VB Scripting\example.txt", True) ' True to overwrite if it exists
file.WriteLine("Hello, World!")
file.Close

' Example: Read from the text file
Set file = fso.OpenTextFile("E:\CS, IOT, Embedded System\Github\VB Scripting\example.txt", 1) ' 1 for reading
Do While Not file.AtEndOfStream
    line = file.ReadLine
    WScript.Echo line
Loop
file.Close

'=========================================================================
' Example: Create a new folder if it doesn't exist
' create a folder
If Not fso.FolderExists("E:\CS, IOT, Embedded System\Github\VB Scripting\MyFolder") Then
    fso.CreateFolder("E:\CS, IOT, Embedded System\Github\VB Scripting\MyFolder")
    WScript.Echo "Folder created."
Else
    WScript.Echo "Folder already exists."
End If

' Example: Copy the text file to a new location
fso.CopyFile "E:\CS, IOT, Embedded System\Github\VB Scripting\example.txt", "E:\CS, IOT, Embedded System\Github\VB Scripting\MyFolder\example_copy.txt"
WScript.Echo "File copied."

' Example: Move the text file to a new location
fso.MoveFile "E:\CS, IOT, Embedded System\Github\VB Scripting\MyFolder\example_copy.txt", "E:\CS, IOT, Embedded System\Github\VB Scripting\MyFolder\example_moved.txt"
WScript.Echo "File moved."

' ============================================================================
' Path of the file to rename
filePath = "E:\CS, IOT, Embedded System\Github\VB Scripting\example.txt"

If fso.FileExists(filePath) Then     ' Check if the file exists
    Set file = fso.GetFile(filePath) ' Get the File object
    file.Name = "renamed_example.txt"   ' Only new name, not full path
    WScript.Echo "File renamed successfully."
Else
    WScript.Echo "File not found."
End If

' Example: Delete the text file
fso.DeleteFile("E:\CS, IOT, Embedded System\Github\VB Scripting\example.txt")
WScript.Echo "File deleted."

' Clean up
Set file = Nothing ' releases memory used by the File object
Set fso = Nothing ' releases memory used by the File System Object
' Note: Be cautious when using file operations, especially delete, as they can permanently remove files from your system.