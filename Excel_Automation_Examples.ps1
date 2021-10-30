# [x]TODO: Create Counter Loop for Filtered Rows.
# [ ]TODO: Create "Log File" function.


# RESOURCES
# How to Automate the opening of an Excel Spreadsheet in Powershell
# https://support.jamsscheduler.com/hc/en-us/articles/206191918-How-to-Automate-the-opening-of-an-Excel-Spreadsheet-in-Powershell

# Scripting::Powershell::5.0:: Working with Microsoft Office Excel - part 1
# https://www.myfaqbase.com/q0001773-Scripting-Powershell-5-0-Working-with-Microsoft-Office-Excel-part-1.html


# Note: The following is necessary for such things as "MessageBox".
# Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Specify the path to the Excel file and the WorkSheet Name
$ExcelFilePath = "C:\Projects\PowerShell\Excel_Automation\Financial Sample.xlsx"

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $true

# OPEN THE EXCEL FILE
$WorkBook = $objExcel.Workbooks.Open($ExcelFilePath)

# SELECT THE WORKSHEET
$WorksheetName = "Sheet1"
# $WorkSheet = $WorkBook.Worksheets.Item($WorksheetName)
$WorkSheet = $WorkBook.Worksheets[$WorksheetName]


# # NUMBER OF WORKSHEET ROWS
$RowCount = ($WorkSheet.UsedRange.Rows).Count
$Notice = "The number of total rows is: '" + $RowCount + "'"
[System.Windows.MessageBox]::Show($Notice)


# # CELL VALUE
# # $CellValue = $WorkSheet.Cells.Item(1, 2).Text
# $CellValue = $WorkSheet.Cells[1, 2].Text
# $Notice = "The cell value is: '" + $CellValue + "'"
# [System.Windows.MessageBox]::Show($Notice)


# MODIFY CELL VALUE
# $WorkSheet.Cells[3, 2].Value = "Columbia"


# APPLY AUTOFILTER
# $WorkSheet.Range("A1:P701").AutoFilter Field:=2, Criteria1:="Germany"
# $WorkSheet.Range("A1:P701").AutoFilter(2,"Germany")
$WorkSheet.UsedRange.AutoFilter(2,"Germany")


# xlCellType ENUERATION (EXCEL)
$xlCellTypeLastCell = 11
$xlCellTypeVisible = 12


# SELECT LAST FILTERED CELL
# $LastCell = $WorkSheet.UsedRange.SpecialCells($xlCellTypeLastCell)
# $LastCellRowIndex = $LastCell.Row
# $LastCell.Select()


# SELECT ALL FILTERED CELLS
# Note: "SpecialCells" returns a non-contiguous range. You must loop through each area.
# $FilteredCells = $WorkSheet.UsedRange.SpecialCells($xlCellTypeVisible)
$FilteredCells = $WorkSheet.AutoFilter.Range.SpecialCells($xlCellTypeVisible)
$FilteredCells.Select()


# LOOP THRU ARRAY
$RowCount = 0
foreach ($FilteredRow in $FilteredCells.Rows) {
    $RowCount++
}


$Notice = "The number of filtered rows is: '" + $RowCount + "'"
[System.Windows.MessageBox]::Show($Notice)



# CLOSE WORKBOOK
$WorkBook.close($false)

# EXIT EXCEL
$objExcel.Quit()