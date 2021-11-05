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

# EXCEL PRIMARY WORKSHEET NAME
# $WorksheetName = "Sheet1"

# EXCEL WORKSHEET OBJECT
$WorkSheet = $null

# HEADER ROW ARRAY
$HeaderRowArray = @()

# COLUMNS IDENTIFIED FOR DELETION
# $DeleteColumnsArray = @("Discount Band", "Units Sold")

# COLUMNS IDENTIFIED FOR RETENTION
# $RetainedHeadersArray = @("Header", "Segment", "Country", "Product", "Date", "Month Number", "Month Name", "Year")
$RetainedHeadersArray = Import-Csv -Path $FilePathRetainedHeaders -Delimiter ","

# COLUMNS TO BE FILTERED
$FieldFiltersArray = Get-Content -Path $FilePathFieldFilters


# FUNCTIONS

#region FUNCTION - FIND HEADER COLUMN INDEX
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
#endregion FUNCTION - FIND HEADER COLUMN INDEX


#region FUNCTION - RETURN MATCHING ARRAY ITEMS
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
#endregion FUNCTION - RETURN MATCHING ARRAY ITEMS


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

# BACKUP WORKBOOK (WITH PASSWORD)
$objExcel.DisplayAlerts = $false; # Note: "$false" = Do not prompt for confirmation to "over-write" backup file.
# $WorkBook.SaveAs($FilePathBackup)
$WorkBook.SaveAs($FilePathModified, [Type]::Missing, $WorkbookPassword)


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


#region CELL VALUE
# $CellValue = $WorkSheet.Cells[1, 2].Text
# $Notice = "The cell value is: '" + $CellValue + "'"
# [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
#endregion CELL VALUE


#region MODIFY CELL VALUE
# $WorkSheet.Cells[3, 2].Value = "Columbia"
#endregion MODIFY CELL VALUE


#region APPLY AUTOFILTER
# READ EACH RECORD INTO ARRAY
# Note: Start with 2nd item in array (i.e. Exclude the column header name).
for ($intRecordCounter=1; $intRecordCounter -lt $FieldFiltersArray.Count; $intRecordCounter++) {
    $Record = ""
    $Record = $FieldFiltersArray[$intRecordCounter]

    # CLEAR AUTOFILTER CRITERIA
    [string[]]$FilterCriteriaArray = @()

    # CLEAR REFINED COLUMN CONTENTS
    $ColumnTextArray = @()

    # CLEAR FILTER FIELD INDEX
    $FilterFieldIndex = 0

    # PARSE EACH ROW ITEM (FROM THE FILTERS FILE)
    $FilterArray = @()
    $FilterArray = $Record.Split(",")

    # REMOVE EMPTY ARRAY ITEMS
    $TrimmedFilterArray = @()
    foreach ($Item in $FilterArray) {
        if (-not ([string]::IsNullOrWhiteSpace($Item))) {
            $TrimmedFilterArray += $Item
        }
    }

    # POPULATE FILTER ARRAY (NO EMPTY ITEMS)
    $FilterArray = @()
    $FilterArray += $TrimmedFilterArray

    # LOOP THROUGH FILTER ARRAY
    for ($intFieldCounter=0; $intFieldCounter -lt $FilterArray.Length; $intFieldCounter++) {
        # GET FILTER ITEM (FROM THE FILTERS FILE)
        $FilterItem = ""
        $FilterItem = $FilterArray[$intFieldCounter] # Example: "Country,Can*,Franc*,Germ*"

        if (-not ([string]::IsNullOrWhiteSpace($FilterItem))) {
            switch ($intFieldCounter) 
            {
                0 {
                    # WORKSHEET EXISTENCE
                    $WorksheetExists = $false

                    # NEW WORKSHEET NAME
                    $NewWorksheetName = $FilterItem

                    # DOES WORKSHEET EXIST?
                    foreach ($WorksheetTab in $WorkBook.Worksheets) {
                        if ($WorksheetTab.Name -eq $NewWorksheetName) {
                            $WorksheetExists = $true
                        }
                    }

                    # ADD NEW WORKSHEET
                    if ($WorksheetExists -eq $false) {
                        # NUMBER OF WORKSHEETS IN WORKBOOK
                        $intWorksheetCount = 0
                        $intWorksheetCount = $WorkBook.Worksheets.Count

                        # ADD NEW WORKSHEET TO WORKBOOK
                        # $WorkBook.Worksheets.Add($NewWorksheetName)
                        # $WorkBook.Worksheets.Add([System.Reflection.Missing]::Value, $WorkBook.Worksheets[$intWorksheetCount], 1, $WorkBook.Worksheets[1].Type)

                        # COPY PRIMARY WORKSHEET (TO LAST POSITION)
                        # $WorkBook.Worksheets[$WorksheetName].Copy([System.Reflection.Missing]::Value, $WorkBook.Worksheets[$intWorksheetCount])
                        $WorkBook.Worksheets[1].Copy([System.Reflection.Missing]::Value, $WorkBook.Worksheets[$intWorksheetCount])

                        # RENAME WORKSHEET (IN LAST POSITION)
                        $WorkBook.Worksheets[$intWorksheetCount + 1].Name = $NewWorksheetName

                        # ACTIVATE WORKSHEET
                        # $WorkBook.Worksheets[$WorksheetName].Activate()
                        $WorkBook.Worksheets[$NewWorksheetName].Activate()

                        # SET REFERENCE TO WORKSHEET OBJECT
                        # $WorkSheet = $WorkBook.Worksheets[$WorksheetName]
                        $WorkSheet = $WorkBook.Worksheets[$NewWorksheetName]

                        #region NUMBER OF ROWS IN WORKSHEET
                        $RowCount = 0
                        $RowCount = $WorkSheet.UsedRange.Rows.Count - 1
                        $Notice = "The total number of rows is: '" + $RowCount + "'"
                        [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
                        #endregion NUMBER OF ROWS IN WORKSHEET

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

                        # GET HEADER ROW ARRAY
                        $HeaderRowArray = $WorkSheet.UsedRange.Rows(1).Cells
                    }
                }

                1 {
                    # CLEAR COLUMN CONTENTS
                    $ColumnContentsArray = @()

                    # FIND HEADER COLUMN INDEX
                    $FilterFieldIndex = FindHeaderColumnIndex $HeaderRowArray $FilterItem # Example: "Country = 2nd column"

                    # FIND COLUMN CONTENTS
                    $ColumnContentsArray = $WorkSheet.UsedRange.Columns($FilterFieldIndex).Cells # Example: 700 rows

                    # POPULATE COLUMN TEXT ARRAY WITH COLUMN CONTENTS (TEXT VALUES)
                    # Note: Start with 2nd item in array (i.e. Exclude the column header name).
                    for ($intCellCounter=2; $intCellCounter -le $ColumnContentsArray.Count; $intCellCounter++) {
                        $ColumnTextArray += $ColumnContentsArray[$intCellCounter].Text
                    }

                    # DE-DUPE AND SORT
                    # Note: Cannot sort "$ColumnContentsArray" as it contains objects and not just text.
                    $ColumnTextArray = $ColumnTextArray | Sort-Object -Unique # Exmple: 8 rows
                }

                default {
                    # SEARCH ARRAY
                    $MatchingArrayItems = @()
                    $MatchingArrayItems = ReturnMatchingArrayItems $ColumnTextArray $FilterItem # Example: Filter = "Can*; Result = "Canada"

                    # ADD TO AUTOFILTER CRITERIA
                    $FilterCriteriaArray += $MatchingArrayItems
                }
            }
        }
        else {
            # ERROR
        }
    }

    # APPLY AUTOFILTER
    $WorkSheet.UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints number of records found.

    #region SELECT ALL FILTERED CELLS
    # Note: "SpecialCells" returns a non-contiguous range. You must loop through each area.
    $FilteredCells = $WorkSheet.AutoFilter.Range.SpecialCells($xlCellTypeVisible)
    $FilteredCells.Select() # Prints "True" if successful.
    #endregion SELECT ALL FILTERED CELLS

    #region NUMBER OF FILTERED WORKSHEET ROWS
    $FilteredRowCount = -1
    foreach ($FilteredRow in $FilteredCells.Rows) {
        $FilteredRowCount++
    }
    $Notice = "The number of filtered rows is: '" + $FilteredRowCount + "'"
    [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
    #endregion NUMBER OF FILTERED WORKSHEET ROWS
}
#endregion APPLY AUTOFILTER


# #region SELECT LAST FILTERED CELL
# # $LastCell = $WorkSheet.UsedRange.SpecialCells($xlCellTypeLastCell)
# # $LastCellRowIndex = $LastCell.Row
# # $LastCellColumnIndex = $LastCell.Column
# # $LastCell.Select()
# #endregion SELECT LAST FILTERED CELL


# CLOSE WORKBOOK
$WorkBook.Close($true) # Note: "$true" = Save file changes. "$false" = Do not save file changes.

# EXIT EXCEL
$objExcel.Quit()