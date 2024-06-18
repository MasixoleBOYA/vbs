'Variables

Dim outputFolder
Dim objExcel
Dim objWorkbook
Dim objWorksheet

'excel file path
outputFolder = "C:\Users\J1121857\Sapworkdir\Customer_List\CustomerList.xlsx"

'Starting excel instance
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

'Opening the workbook
Set objWorkbook = objExcel.Workbooks.Open(outputFolder)

If objWorkbook Is Nothing Then
	MsgBox"Workbook does not exist"
Else
	MsgBox"Workbook exists"
End If

'Opening the worksheet
Set objWorksheet = objWorkbook.Worksheets(1)

If objWorksheet Is Nothing Then
	MsgBox"Worksheet could not open"
	WScript.Quit
Else
	MsgBox"Worksheet opened successfully"
End If

'Peform the operations
Dim objRange
Set objRange = objWorksheet.UsedRange


'Dim objTarget 
'Set objTarget = objRange.Find("Name")
'MsgBox "Your column index is"& objTarget