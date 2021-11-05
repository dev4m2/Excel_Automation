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
$FileDirectory = "C:\Projects\PowerShell\Excel_Automation\"
$Filename = "Financial Sample"
$FilePath = $FileDirectory + $Filename + "." + "xlsx"
$FilePathBackup = $FileDirectory + $Filename + ".bak" + "." + "xlsx"
$FilePathModified = $FileDirectory + $Filename + " " + $Date + "." + "xlsx"
$FilePathRetainedHeaders = $FileDirectory + $Filename + " - Headers" + "." + "csv"
$FilePathFieldFilters = $FileDirectory + $Filename + " - Filters" + "." + "csv"

# EXCEL WORKBOOK PASSWORD
$WorkbookPassword = "password"

# EXCEL WORKSHEET
$WorksheetName = "Sheet1"

# NEW EXCEL WORKSHEET
$WorksheetModifiedName = "Modified"

# COLUMNS IDENTIFIED FOR DELETION
# $DeleteColumnsArray = @("Discount Band", "Units Sold")

# COLUMNS IDENTIFIED FOR RETENTION
# $RetainedHeadersArray = @("Segment", "Country", "Product", "Date", "Month Number", "Month Name", "Year")
$RetainedHeadersArray = Import-Csv -Path $FilePathRetainedHeaders -Delimiter ","

# COLUMNS TO BE FILTERED
# $FieldFiltersArray = Import-Csv -Path $FilePathFieldFilters -Delimiter ","
$FieldFiltersArray = Get-Content -Path $FilePathFieldFilters


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


#region RETURN MATCHING ARRAY ITEMS
function ReturnMatchingArrayItems {
    # param ($refArray, $refSearchString)
    param ([string[]]$refArray, [string]$refSearchString)

    $refMatchingItemsArray = @()
    foreach ($refTextItem in $refArray) {
        if ($refTextItem -Like $refSearchString) {
            $refMatchingItemsArray += $refTextItem
        }
    }
    return $refMatchingItemsArray
}
#endregion RETURN MATCHING ARRAY ITEMS


#region RETAIN SPECIFIED COLUMNS
# function RetainSpecifiedColumns {
#     param ()
# }
#endregion RETAIN SPECIFIED COLUMNS


# BACKUP FILE(S)
# Copy-Item $FilePath -Destination $FilePathBackup -Force
Copy-Item $FilePath -Destination $FilePathModified -Force

# Note: The following is necessary for such things as "MessageBox".
# Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $true


# OPEN THE EXCEL FILE
# $WorkBook = $objExcel.Workbooks.Open($FilePath)
$WorkBook = $objExcel.Workbooks.Open($FilePathModified)


# $TestWorksheetCount = $WorkBook.Worksheets.Count


# BACKUP WORKBOOK (WITH PASSWORD)
$objExcel.DisplayAlerts = $false; # Note: "$false" = Do not prompt for confirmation to "over-write" backup file.
# $WorkBook.SaveAs($FilePathBackup)
$WorkBook.SaveAs($FilePathModified, [Type]::Missing, $WorkbookPassword)


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
# # Create array of column names to delete
# foreach ($FieldName in $DeleteColumnsArray) {
#     $ColumnIndex = 0

#     # GET HEADER ROW
#     $AdjustedHeaderRowArray = @() # Clear array
#     $AdjustedHeaderRowArray = $WorkSheet.UsedRange.Rows(1).Cells

#     foreach ($Column in $AdjustedHeaderRowArray) {
#         $ColumnIndex++
#         $ColumnText = $Column.Text
#         if ($ColumnText -eq $FieldName) {
#             $WorkSheet.Columns[$ColumnIndex].EntireColumn.Delete() # Prints "True" if successful.
#             break
#         }
#     }
# }
#endregion DELETE SPECIFIED COLUMNS


#region RETAIN SPECIFIED COLUMNS
$ColumnIndex = 1
$FieldCount = $WorkSheet.UsedRange.Rows(1).Cells.Count

while ($ColumnIndex -le $FieldCount) {
    $Column = $WorkSheet.UsedRange.Rows(1).Cells[$ColumnIndex]
    $ColumnText = $Column.Text

    # if ($RetainedHeadersArray -match $ColumnText) {
    if ($RetainedHeadersArray.Header -match $ColumnText) {
        # $ResponseText = "Field Header: '" + $ColumnText + "' should be retained."
        $ColumnIndex++
    }
    else {
        # $ResponseText = "Field Header: '" + $ColumnText + "' should be DELETED."
        $WorkSheet.Columns[$ColumnIndex].EntireColumn.Delete() # Prints "True" if successful.
        $FieldCount = $WorkSheet.UsedRange.Rows(1).Cells.Count
    }

    # Write-Output $ResponseText
}
#endregion RETAIN SPECIFIED COLUMNS


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


#region GET HEADER ROW ARRAY
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
foreach ($Record in $FieldFiltersArray) {
    # CLEAR AUTOFILTER CRITERIA
    [string[]]$FilterCriteriaArray = @()

    # CLEAR COLUMN CONTENTS
    $ColumnContentsArray = @()

    $FilterArray = $Record.Split(",")
    for ($intCriteriaCounter=0; $intCriteriaCounter -lt $FilterArray.Length; $intCriteriaCounter++) {
        $FilterItem = $FilterArray[$intCriteriaCounter]
        if (-not ([string]::IsNullOrWhiteSpace($FilterItem))) {
            if ($intCriteriaCounter -eq 0) {
                # FIND HEADER COLUMN INDEX
                $FilterFieldName = $FilterItem
                $FilterFieldIndex = FindHeaderColumnIndex $HeaderRowArray $FilterFieldName

                # FIND COLUMN CONTENTS
                $ColumnContentsArray = @()
                $ColumnContentsArray = $WorkSheet.UsedRange.Columns($FilterFieldIndex).Cells
                # foreach ($Cell in $WorkSheet.UsedRange.Columns($FilterFieldIndex).Cells) {
                #     $ColumnContentsArray += $Cell.Text
                # }

                $ColumnTextArray = @()
                # $ColumnContentsTotalCount = $ColumnContentsArray.Count
                for ($intCellCounter=2; $intCellCounter -le $ColumnContentsArray.Count; $intCellCounter++) {
                    $ColumnTextArray += $ColumnContentsArray[$intCellCounter].Text
                }

                # DE-DUPE AND SORT
                $ColumnTextArray = $ColumnTextArray | Sort-Object -Unique
            }
            else {
                # SEARCH ARRAY
                $SearchString = $FilterItem # Example: "Can*"
                $MatchingArrayItems = @()
                # $MatchingArrayItems = ReturnMatchingArrayItems $ColumnContentsArray $SearchString
                $MatchingArrayItems = ReturnMatchingArrayItems $ColumnTextArray $SearchString

                # DE-DUPE AND SORT
                # $MatchingArrayItems = $MatchingArrayItems | Sort-Object -Unique

                # ADD TO AUTOFILTER CRITERIA
                $FilterCriteriaArray += $MatchingArrayItems
            }
        }
        else {
            # ERROR
        }
    }

    # APPLY AUTOFILTER
    $WorkSheet.UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints number of records found.
}
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