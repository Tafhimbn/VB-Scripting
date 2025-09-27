' Conditional Statements in VBScript
' ============================================================================
Dim x, y
x = 10
y = 20
'Conditional statements used in VBScript:
'1) If...Then → Runs code if condition is true.
If x < y Then
    MsgBox "x is less than y"
End If

' ii) If...Then...Else → Chooses between two paths.
If x > y Then
    MsgBox "x is greater than y"
Else
    MsgBox "x is not greater than y"
End If

' iii) If...ElseIf...Else → Handles multiple conditions.
If x < y Then
    MsgBox "x is less than y"
ElseIf x > y Then
    MsgBox "x is greater than y"
Else
    MsgBox "x is equal to y"
End If

' iv) Select Case → Cleaner alternative for multiple If...ElseIf.
Select Case x
    Case 10
        MsgBox "x is exactly 10"
    Case 20
        MsgBox "x is exactly 20"
    Case Else
        MsgBox "x is something else"
End Select
'In VBScript, there is no Switch statement like in C/C++/JavaScript.
'Instead, VBScript uses Select Case, which works the same way as a switch-case.