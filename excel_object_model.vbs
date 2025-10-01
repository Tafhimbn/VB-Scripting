' Excel Object Model
' ============================================================================
' The Excel Object Model is a hierarchical structure that represents all the
' objects in an Excel application. It allows you to manipulate Excel files
' programmatically using VBScript.
'
' Key Objects in the Excel Object Model:
' 1. Application: Represents the entire Excel application.
' 2. Workbook: Represents an Excel workbook (file).
' 3. Worksheet: Represents a single worksheet within a workbook.
' 4. Range: Represents a cell or a range of cells in a worksheet.
' ============================================================================
' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")

    ' CreateObject("Excel.Application") → Creates a new instance of Microsoft Excel.
    ' Set excelApp = ... → Stores that Excel instance in the variable excelApp.
    ' At this point, Excel is running in the background (invisible by default).

excelApp.Visible = True ' Make Excel visible

excelApp.DisplayAlerts = False ' Disable alerts

Set workbook = excelApp.Workbooks.Add() ' Add a new workbook

Set worksheet = workbook.Worksheets.Add() ' Add a new worksheet

worksheet.Name = "MySheet" ' Rename the worksheet

' Set values in specific cells
worksheet.Range("A1").Value = "Hello"
worksheet.Range("B1").Value = "World"

' Set a formula in a cell
worksheet.Range("C1").Formula = "=A1 & "" "" & B1"

' Autofit columns to the content
worksheet.Columns("A:C").AutoFit

' Save the workbook
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx"

workbook.Close False ' Close the workbook without saving changes

excelApp.Quit
' Clean up
Set worksheet = Nothing ' releases memory used by the Worksheet object
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object


