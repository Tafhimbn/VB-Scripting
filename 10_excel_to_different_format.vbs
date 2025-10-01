' Excel to PDF using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application") 
excelApp.Visible = False                         ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False                    ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")

' Export the workbook to PDF format
workbook.ExportAsFixedFormat 0, "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.pdf"

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to CSV using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False                          ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False                    ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")

' Save the workbook as CSV format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.csv", 6 ' 6 is the file format for CSV

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to TXT using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False       ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")

' Save the workbook as TXT format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.txt", 42 ' 42 is the file format for TXT

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to HTML using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False       ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")

' Save the workbook as HTML format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.html", 44 ' 44 is the file format for HTML

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to XML using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")

excelApp.Visible = False ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")

' Save the workbook as XML format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xml", 46 ' 46 is the file format for XML

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to ODS using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False       ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")
' Save the workbook as ODS format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.ods", 60 ' 60 is the file format for ODS

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to XLS using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False       ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False ' Disable alerts
' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")
' Save the workbook as XLS format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xls", 56 ' 56 is the file format for XLS

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to XLSM using VBScript
' ============================================================================

' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False       ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")
' Save the workbook as XLSM format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsm", 52 ' 52 is the file format for XLSM

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object

' Excel to XLSB using VBScript
' ============================================================================  
    
' Create an instance of the Excel application
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False       ' Keep Excel hidden during the process
excelApp.DisplayAlerts = False ' Disable alerts

' Open the existing Excel workbook
Set workbook = excelApp.Workbooks.Open("E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsx")
' Save the workbook as XLSB format
workbook.SaveAs "E:\CS, IOT, Embedded System\Github\VB Scripting\MyExcelFile.xlsb", 50 ' 50 is the file format for XLSB

' Close the workbook without saving changes
workbook.Close False
excelApp.Quit

' Clean up
Set workbook = Nothing ' releases memory used by the Workbook object
Set excelApp = Nothing ' releases memory used by the Excel Application object


