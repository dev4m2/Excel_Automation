# [x]TODO: Create Counter Loop for Filtered Rows.
# [ ]TODO: Create "Log File" function.
# [ ]TODO: Create "FindHeaderColumnIndex" function.


# RESOURCES
# How to Automate the opening of an Excel Spreadsheet in Powershell
# https://support.jamsscheduler.com/hc/en-us/articles/206191918-How-to-Automate-the-opening-of-an-Excel-Spreadsheet-in-Powershell

# Scripting::Powershell::5.0:: Working with Microsoft Office Excel - part 1
# https://www.myfaqbase.com/q0001773-Scripting-Powershell-5-0-Working-with-Microsoft-Office-Excel-part-1.html


# GLOBAL VARIABLES

# OPERATOR PARMETER CONSTANTS
$xlAnd = 1
$xlOr  = 2
$xlTop10Items = 3
$xlBottom10Items = 4
$xlTop10Percent = 5
$xlBottom10Percent = 6
$xlFilterValues = 7
$xlFilterCellColor = 8
$xlFilterFontColor = 9
$xlFilterIcon = 10
$xlFilterDynamic = 11


# xlCellType ENUMERATION CONSTANTS (EXCEL)
$xlCellTypeLastCell = 11
$xlCellTypeVisible = 12

# DATE INFO
$Date = Get-Date -Format "yyyy-MM-dd"

# EXCEL FILEPATH
$Filename = "C:\Projects\PowerShell\Excel_Automation\Financial Sample"
$FilenameExtension = "xlsx"
$FilePath = $Filename + "." + $FilenameExtension
$FileBackupPath = $Filename + ".bak" + "." + $FilenameExtension
$FileModifiedPath = $Filename + " " + $Date + "." + $FilenameExtension

# EXCEL WORKSHEET
$WorksheetName = "Sheet1"

# NEW EXCEL WORKSHEET
$WorksheetModifiedName = "Modified"

# COLUMNS IDENTIFIED FOR DELETION
$DeleteColumnsArray = @("Discount Band", "Units Sold")


# FUNCTIONS

#region FIND HEADER COLUMN INDEX
function FindHeaderColumnIndex {
    param ($refHeaderRowArray, $refFilterFieldName)

    # $test1 = $refHeaderRowArray.GetType()
    # $test2 = $refFilterFieldName.GetType()

    $ColumnIndex = 0
    $IndexResult = 0
    foreach ($Column in $refHeaderRowArray) {
        $ColumnIndex++
        $ColumnText = $Column.Text
        if ($refFilterFieldName -eq $ColumnText) {
            $IndexResult = $ColumnIndex
            break
        }
    }
    return $IndexResult
}
#endregion FIND HEADER COLUMN INDEX


# BACKUP FILE(S)
# Copy-Item $FilePath -Destination $FileBackupPath -Force
Copy-Item $FilePath -Destination $FileModifiedPath -Force

# Note: The following is necessary for such things as "MessageBox".
# Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $true


# OPEN THE EXCEL FILE
# $WorkBook = $objExcel.Workbooks.Open($FilePath)
$WorkBook = $objExcel.Workbooks.Open($FileModifiedPath)


# BACKUP WORKBOOK (WITH PASSWORD)
$objExcel.DisplayAlerts = $false; # Note: "$false" = Do not prompt for confirmation to "over-write" backup file.
# $WorkBook.SaveAs($FileBackupPath)
$WorkBook.SaveAs($FileModifiedPath, [Type]::Missing, "password")


# ADD NEW WORKSHEET TO WORKBOOK
# $WorkBook.Worksheets.Add($WorksheetModifiedName)
# $WorkBook.Worksheets.Add([System.Reflection.Missing]::Value, $WorkBook.Worksheets[$WorkBook.Worksheets.Count], 1, $WorkBook.Worksheets[1].Type)


# COPY WORKSHEET
$WorkBook.Worksheets[$WorksheetName].Copy([System.Reflection.Missing]::Value, $WorkBook.Worksheets[$WorkBook.Worksheets.Count])


# RENAME WORKSHEET
$WorkBook.Worksheets[$WorkBook.Worksheets.Count].Name = $WorksheetModifiedName


# ACTIVATE WORKSHEET
$WorkBook.Worksheets[$WorksheetName].Activate()


# SET REFERENCE TO WORKSHEET OBJECT
$WorkSheet = $WorkBook.Worksheets[$WorksheetName]


#region DELETE SPECIFIED COLUMNS
# Create array of column names to delete
foreach ($FieldName in $DeleteColumnsArray) {
    $ColumnIndex = 0

    # GET HEADER ROW
    $AdjustedHeaderRowArray = @() # Clear array
    $AdjustedHeaderRowArray = $WorkSheet.UsedRange.Rows(1).Cells

    foreach ($Column in $AdjustedHeaderRowArray) {
        $ColumnIndex++
        $ColumnText = $Column.Text
        if ($ColumnText -eq $FieldName) {
            $WorkSheet.Columns[$ColumnIndex].EntireColumn.Delete() # Prints "True" if successful.
            break
        }
    }
}
#endregion DELETE SPECIFIED COLUMNS


#region CELL VALUE
# $CellValue = $WorkSheet.Cells[1, 2].Text
# $Notice = "The cell value is: '" + $CellValue + "'"
# [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
#endregion CELL VALUE


#region MODIFY CELL VALUE
# $WorkSheet.Cells[3, 2].Value = "Columbia"
#endregion MODIFY CELL VALUE


#region NUMBER OF WORKSHEET ROWS
$RowCount = $WorkSheet.UsedRange.Rows.Count
$Notice = "The total number of rows is: '" + $RowCount + "'"
[System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
#endregion NUMBER OF WORKSHEET ROWS


#region GET HEADER ROW
$HeaderRowArray = @() # Clear array
$HeaderRowArray = $WorkSheet.UsedRange.Rows(1).Cells
#endregion GET HEADER ROW


#region FIND HEADER COLUMN COUNT
# $ColumnCount = 0
# $ColumnCount = $HeaderRowArray.Count
# $Notice = "The total number of columns is: '" + $ColumnCount + "'"
# [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
#endregion FIND HEADER COLUMN COUNT


#region APPLY AUTOFILTER
# FIND HEADER COLUMN INDEX
$FilterFieldName = "Country"
$FilterFieldIndex = FindHeaderColumnIndex $HeaderRowArray $FilterFieldName


# AUTOFILTER CRITERIA
$FilterCriteriaArray = @()
$FilterCriteriaArray += "=France"
$FilterCriteriaArray += "=Germany"
$WorkSheet.UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints number of records found.
#endregion APPLY AUTOFILTER


#region APPLY AUTOFILTER
# FIND HEADER COLUMN INDEX
$FilterFieldName = "Product"
$FilterFieldIndex = FindHeaderColumnIndex $HeaderRowArray $FilterFieldName


# AUTOFILTER CRITERIA
$FilterCriteriaArray = @()
$FilterCriteriaArray += "=Paseo"
$FilterCriteriaArray += "=Velo"
$WorkSheet.UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints number of records found.
#endregion APPLY AUTOFILTER


#region APPLY AUTOFILTER
# FIND HEADER COLUMN INDEX
$FilterFieldName = "Month Number"
$FilterFieldIndex = FindHeaderColumnIndex $HeaderRowArray $FilterFieldName


# AUTOFILTER CRITERIA
$FilterCriteriaArray = @()
$FilterCriteriaArray += "=6"
$FilterCriteriaArray += "=10"
$WorkSheet.UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints number of records found.
#endregion APPLY AUTOFILTER


#region APPLY AUTOFILTER
# FIND HEADER COLUMN INDEX
$FilterFieldName = "Year"
$FilterFieldIndex = FindHeaderColumnIndex $HeaderRowArray $FilterFieldName


# AUTOFILTER CRITERIA
$FilterCriteriaArray = @()
$FilterCriteriaArray += "=2013*"
$FilterCriteriaArray += "=2014"
$WorkSheet.UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints number of records found.
#endregion APPLY AUTOFILTER


#region SELECT ALL FILTERED CELLS
# Note: "SpecialCells" returns a non-contiguous range. You must loop through each area.
# $FilteredCells = $WorkSheet.UsedRange.SpecialCells($xlCellTypeVisible)
$FilteredCells = $WorkSheet.AutoFilter.Range.SpecialCells($xlCellTypeVisible)
$FilteredCells.Select() # Prints "True" if successful.
#endregion SELECT ALL FILTERED CELLS


#region SELECT LAST FILTERED CELL
# $LastCell = $WorkSheet.UsedRange.SpecialCells($xlCellTypeLastCell)
# $LastCellRowIndex = $LastCell.Row
# $LastCellColumnIndex = $LastCell.Column
# $LastCell.Select()
#endregion SELECT LAST FILTERED CELL


#region NUMBER OF FILTERED WORKSHEET ROWS
$RowCount = 0
foreach ($FilteredRow in $FilteredCells.Rows) {
    $RowCount++
}
$Notice = "The number of filtered rows is: '" + $RowCount + "'"
[System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
#endregion NUMBER OF FILTERED WORKSHEET ROWS

# CLOSE WORKBOOK
$WorkBook.Close($true) # Note: "$true" = Save file changes. "$false" = Do not save file changes.

# EXIT EXCEL
$objExcel.Quit()