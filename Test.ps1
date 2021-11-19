# [x]TODO: Create Counter Loop for Filtered Rows.
# [ ]TODO: Create "Log File" function.
# [x]TODO: Adjust for optional password.
# [ ]TODO: Allow wildcard (i.e. "*") for headers retention.
# [ ]TODO: Allow criteria filtering on "blanks".


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
$WorkbookPassword = ""

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

#region FUNCTION - RETURN MATCHING ARRAY ITEMS
function ReturnMatchingArrayItems {
    # param ($refArray, $refSearchString)
    param ([string[]]$refArray, [string]$refSearchString)

    $refMatchingItemsArray = @()
    foreach ($refTextItem in $refArray) {
        if ($refSearchString -eq '""') { # Is search string empty?
            if ([string]::IsNullOrWhiteSpace($refTextItem)) { # Filter on blank cells.
                $refMatchingItemsArray += "="
                break
            }
        }
        elseif ($refTextItem -Like $refSearchString) { # Filter on wildcards.
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

# ACTIVATE WORKSHEET
$WorkBook.Worksheets[1].Activate()

# SET REFERENCE TO WORKSHEET OBJECT
$WorkSheet = $WorkBook.Worksheets[1]


#region CELL VALUE
# $CellValue = $WorkSheet.Cells[1, 2].Text
# $Notice = "Cell value: '" + $CellValue + "'"
# [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
#endregion CELL VALUE


#region MODIFY CELL VALUE
# $WorkSheet.Cells[3, 2].Value = "Columbia"
#endregion MODIFY CELL VALUE

$ColumnTextArray = @()
# $ColumnTextArray = $WorkSheet.UsedRange.Rows.Columns[13].Cells | ForEach-Object { $_.Text }

# $ColumnTextArray = $WorkSheet.Range("M2", "M50").Cells
$ColumnTextArray = $WorkSheet.Rows.Columns[13]


# CLOSE WORKBOOK
$WorkBook.Close($true) # Note: "$true" = Save file changes. "$false" = Do not save file changes.

# EXIT EXCEL
$objExcel.Quit()

# RELEASE RESOURCES
While([System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) -ge 0){}
while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) -ge 0){}
