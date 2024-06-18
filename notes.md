# Explanation of VBScript to Filter and Delete Rows in Excel

## Variable Declarations and Constants
- `Dim outputFolder`: Declares a variable to hold the path to the Excel file.
- `Dim objExcel`: Declares a variable for the Excel application object.
- `Dim objWorkbook`: Declares a variable for the workbook object.
- `Dim objWorksheet`: Declares a variable for the worksheet object.
- `Dim rng`: Declares a variable for the range object.
- `Dim colIndex`: Declares a variable for the column index.

## Constants
- `outputFolder = "C:\Users\J1121857\Sapworkdir\PowerBI_input\yourfile.xlsx"`: Sets the path to the Excel file.

## Create Excel Application Object
- `Set objExcel = CreateObject("Excel.Application")`: Creates an instance of the Excel application.
- `objExcel.Visible = True`: Makes the Excel application visible to the user.

## Open the Workbook
- `Set objWorkbook = objExcel.Workbooks.Open(outputFolder)`: Opens the specified workbook.
- `If objWorkbook Is Nothing Then`: Checks if the workbook was opened successfully.
  - `MsgBox "Error opening workbook: " & outputFolder, vbCritical`: Displays an error message if the workbook failed to open.
  - `WScript.Quit`: Exits the script.

## Access the First Worksheet
- `Set objWorksheet = objWorkbook.Worksheets(1)`: Accesses the first worksheet in the workbook.
- `If objWorksheet Is Nothing Then`: Checks if the worksheet was accessed successfully.
  - `MsgBox "Error accessing worksheet in workbook: " & outputFolder, vbCritical`: Displays an error message if the worksheet failed to access.
  - `objWorkbook.Close False`: Closes the workbook without saving changes.
  - `objExcel.Quit`: Quits the Excel application.
  - `WScript.Quit`: Exits the script.

## Ensure the Worksheet Exists and Apply Filters
- `If Not objWorksheet Is Nothing Then`: Ensures the worksheet object is valid.

### Step 1: Filter and Delete Rows for "Profit centre" = "ZA11232000"
- `Set rng = objWorksheet.Range("A1").CurrentRegion`: Sets the range to the current region starting from cell A1.
- `colIndex = objWorksheet.Range("A1:Z1").Find("Profit centre").Column`: Finds the column index for "Profit centre".
- `rng.AutoFilter Field:=colIndex, Criteria1:="<>ZA11232000"`: Applies a filter to exclude "ZA11232000".
- `objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete`: Deletes the visible rows that do not match the filter.
- `rng.AutoFilter`: Clears the filter.

### Step 2: Filter and Delete Rows for "Name" Starting with 'INACTIVE', 'Inactive', 'CLOSED', 'Closed'
- `colIndex = objWorksheet.Range("A1:Z1").Find("Name").Column`: Finds the column index for "Name".
- `rng.AutoFilter Field:=colIndex, Criteria1:="=INACTIVE*", Operator:=1`: Applies a filter for names starting with 'INACTIVE'.
- `rng.AutoFilter Field:=colIndex, Criteria1:="=Inactive*", Operator:=1, Criteria2:="=CLOSED*", Operator:=1, Criteria3:="=Closed*"`: Extends the filter for names starting with 'Inactive', 'CLOSED', or 'Closed'.
- `objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete`: Deletes the visible rows that match the filter.
- `rng.AutoFilter`: Clears the filter.

### Step 3: Filter and Delete Rows for "SOff.Description" = "Reseller SaleManagSA"
- `colIndex = objWorksheet.Range("A1:Z1").Find("SOff.Description").Column`: Finds the column index for "SOff.Description".
- `rng.AutoFilter Field:=colIndex, Criteria1:="<>Reseller SaleManagSA"`: Applies a filter to exclude "Reseller SaleManagSA".
- `objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete`: Deletes the visible rows that do not match the filter.
- `rng.AutoFilter`: Clears the filter.

### Step 4: Filter and Delete Rows for "Account Group" = "Z001"
- `colIndex = objWorksheet.Range("A1:Z1").Find("Account Group").Column`: Finds the column index for "Account Group".
- `rng.AutoFilter Field:=colIndex, Criteria1:="<>Z001"`: Applies a filter to exclude "Z001".
- `objWorksheet.UsedRange.SpecialCells(12).EntireRow.Delete`: Deletes the visible rows that do not match the filter.
- `rng.AutoFilter`: Clears the filter.

### Save the Workbook
- `objWorkbook.Save`: Saves the workbook.

### Inform the User
- `MsgBox "Workbook processed and saved successfully!", vbInformation`: Displays a success message.

### Error Handling for Missing Worksheet
- `Else`: Handles the case where the worksheet is not found.
  - `MsgBox "Worksheet not found!", vbCritical`: Displays an error message if the worksheet is missing.

## Close the Workbook and Quit Excel
- `objWorkbook.Close False`: Closes the workbook without saving any further changes.
- `objExcel.Quit`: Quits the Excel application.

## Clean Up
- `Set objWorksheet = Nothing`: Releases the worksheet object.
- `Set objWorkbook = Nothing`: Releases the workbook object.
- `Set objExcel = Nothing`: Releases the Excel application object.
