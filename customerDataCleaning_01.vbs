' Declare variables
Dim outputFolder
Dim objExcel
Dim objWorkbook
Dim objWorksheet
Dim rng
Dim colIndex

' Constants
outputFolder = "C:\Users\J1121857\Sapworkdir\PowerBI_input\yourfile.xlsx"

' Create an Excel application object
Set objExcel = CreateObject("Excel.Application")

' Make Excel visible (False to run in the background)
objExcel.Visible = True

' Open the workbook
Set objWorkbook = objExcel.Workbooks.Open(outputFolder)

' Check if the workbook was opened successfully
If objWorkbook Is Nothing Then
    MsgBox "Error opening workbook: " & outputFolder, vbCritical
    WScript.Quit
End If

' Access the first worksheet
Set objWorksheet = objWorkbook.Worksheets(1)

' Check if the worksheet was accessed successfully
If objWorksheet Is Nothing Then
    MsgBox "Error accessing worksheet in workbook: " & outputFolder, vbCritical
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit
End If

' Ensure the worksheet exists
If Not objWorksheet Is Nothing Then
    ' Step 1: Filter and delete rows for "Profit centre" = "ZA11232000"
    Set rng = objWorksheet.Range("A1").CurrentRegion
    colIndex = objWorksheet.Range("A1:Z1").Find("Profit centre").Column
    rng.AutoFilter Field:=colIndex, Criteria1:="<>ZA11232000"
    objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete ' xlCellTypeVisible = 12
    rng.AutoFilter

    ' Step 2: Filter and delete rows for "Name" starting with 'INACTIVE', 'Inactive', 'CLOSED', 'Closed'
    colIndex = objWorksheet.Range("A1:Z1").Find("Name").Column
    rng.AutoFilter Field:=colIndex, Criteria1:="=INACTIVE*", Operator:=1
    rng.AutoFilter Field:=colIndex, Criteria1:="=Inactive*", Operator:=1, Criteria2:="=CLOSED*", Operator:=1, Criteria3:="=Closed*"
    objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete ' xlCellTypeVisible = 12
    rng.AutoFilter

    ' Step 3: Filter and delete rows for "SOff.Description" = "Reseller SaleManagSA"
    colIndex = objWorksheet.Range("A1:Z1").Find("SOff.Description").Column
    rng.AutoFilter Field:=colIndex, Criteria1:="<>Reseller SaleManagSA"
    objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete ' xlCellTypeVisible = 12
    rng.AutoFilter

    ' Step 4: Filter and delete rows for "Account Group" = "Z001"
    colIndex = objWorksheet.Range("A1:Z1").Find("Account Group").Column
    rng.AutoFilter Field:=colIndex, Criteria1:="<>Z001"
    objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete ' xlCellTypeVisible = 12
    rng.AutoFilter

    ' Save the workbook
    objWorkbook.Save

    ' Inform the user
    MsgBox "Workbook processed and saved successfully!", vbInformation
Else
    MsgBox "Worksheet not found!", vbCritical
End If

' Close the workbook and quit Excel
objWorkbook.Close False
objExcel.Quit

' Clean up
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
