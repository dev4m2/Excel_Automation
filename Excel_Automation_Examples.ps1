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


# CELL COLORS
$ColorTable = @{
    Black = 1;
    White = 2;
    Red = 3;
    Lime = 4;
    Blue = 5;
    Yellow = 6;
    Fuchsia = 7;
    Aqua = 8;
    Maroon = 9;
    Green = 10;
    Navy = 11;
    Olive = 12;
    Purple = 13;
    Teal = 14;
    Silver = 15;
    Gray = 16
}


# FUNCTIONS
function ReturnMatchingArrayItems {
    # param ($refArray, $refSearchString)
    param ([string[]]$refArray, [string]$refSearchString)

    $refMatchingItemsArray = @()
    
    foreach ($refTextItem in $refArray) {
        # if ($refSearchString -eq '""') { # Is search string empty?
        if ($refSearchString -eq '""""') { # Is search string empty? (Note: This works for both Excel and Notepad.)
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


# DATE INFO
$Date = Get-Date -Format "yyyy-MM-dd"

# EXCEL FILEPATH
$FileDirectory = "C:\Projects\PowerShell\Excel_Automation\"
$Filename = "Financial Sample"
$ModifiedFilename = "Financial Sample"
# $ModifiedFilename = "Modified"
$FilePath = $FileDirectory + $Filename + "." + "xlsx"
# $FilePathBackup = $FileDirectory + $Filename + ".bak" + "." + "xlsx"
$FilePathModified = $FileDirectory + $ModifiedFilename + " " + $Date + "." + "xlsx"
$FilePathRetainedHeaders = $FileDirectory + $Filename + " - Headers" + "." + "csv"
# $FilePathActionItems = $FileDirectory + $Filename + " - Filters" + "." + "csv"
$FilePathActionItems = $FileDirectory + $Filename + " - Actions" + "." + "csv"

# EXCEL WORKBOOK PASSWORD
$WorkbookPassword = ""

# EXCEL WORKSHEET OBJECT
$WorkSheet = $null

# WORKSHEET NAME
$PreviousWorksheetName = ""

# FILTERED CELL ARRAY
$FilteredCells = $null

# COLUMNS IDENTIFIED FOR DELETION
# $DeleteColumnsArray = @("Discount Band", "Units Sold")

# COLUMNS IDENTIFIED FOR RETENTION
# $RetainedHeadersArray = @("Header", "Segment", "Country", "Product", "Date", "Month Number", "Month Name", "Year")
$RetainedHeadersArray = Import-Csv -Path $FilePathRetainedHeaders -Delimiter ","

# COLUMNS TO BE FILTERED
$ActionItemsArray = Get-Content -Path $FilePathActionItems


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
if (-not ([string]::IsNullOrWhiteSpace($WorkbookPassword))) {
    $objExcel.DisplayAlerts = $false; # Note: "$false" = Do not prompt for confirmation to "over-write" backup file.
    $WorkBook.SaveAs($FilePathModified, [Type]::Missing, $WorkbookPassword)
    $objExcel.DisplayAlerts = $true; # Note: "$true" = Prompt for confirmation to "over-write" backup file.
}


#region CELL VALUE
# $CellValue = $WorkSheet.Cells[1, 2].Text
# $Notice = "Cell value: '" + $CellValue + "'"
# [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
#endregion CELL VALUE


#region MODIFY CELL VALUE
# $WorkSheet.Cells[3, 2].Value = "Columbia"
#endregion MODIFY CELL VALUE


#region LOOP THROUGH ACTION ITEMS
# READ EACH RECORD INTO ARRAY
# Note: Start with 2nd item in array (i.e. Exclude the column header name).
for ($intRecordCounter=1; $intRecordCounter -lt $ActionItemsArray.Count; $intRecordCounter++) {
    $ActionRecord = ""
    $ActionRecord = $ActionItemsArray[$intRecordCounter]

    # CLEAR AUTOFILTER CRITERIA
    [string[]]$FilterCriteriaArray = @()

    # CLEAR ACTION
    $Action = ""

    # CLEAR REFINED COLUMN CONTENTS
    $ColumnTextArray = @()

    # CLEAR FILTER FIELD INDEX
    $FilterFieldIndex = 0

    # PARSE EACH ROW ITEM (FROM THE FILTERS FILE)
    $ActionArray = @()
    $ActionArray = $ActionRecord.Split(',')

    # REMOVE EMPTY ARRAY ITEMS
    $TrimmedFilterArray = @()
    foreach ($Item in $ActionArray) {
        if (-not ([string]::IsNullOrWhiteSpace($Item))) {
            $TrimmedFilterArray += $Item
        }
    }

    # POPULATE ACTION ARRAY (NO EMPTY ITEMS)
    $ActionArray = @()
    $ActionArray += $TrimmedFilterArray

    # WORKSHEET NAME
    $WorksheetName = ""

    # LOOP THROUGH ACTION ITEM ARRAY
    for ($intFieldCounter=0; $intFieldCounter -lt $ActionArray.Length; $intFieldCounter++) {
        # GET ACTION ITEM (FROM THE FILTERS FILE)
        $ActionItem = ""
        $ActionItem = $ActionArray[$intFieldCounter] # Example: "Country,Can*,Franc*,Germ*"

        switch ($intFieldCounter) 
        {
            0 { # WORKSHEET
                if ($PreviousWorksheetName -ne $ActionItem) {
                    $WorksheetExists = $false

                    # WORKSHEET NAME
                    # $WorksheetName = ""
                    $WorksheetName = $ActionItem
                    $PreviousWorksheetName = $ActionItem
    
                    # DOES WORKSHEET EXIST?
                    foreach ($WorksheetTab in $WorkBook.Worksheets) {
                        if ($WorksheetTab.Name -eq $WorksheetName) {
                            $WorksheetExists = $true
                            break
                        }
                    }
    
                    # SELECT EXISTING WORKSHEET
                    if ($WorksheetExists -eq $true) {
                        # ACTIVATE WORKSHEET
                        $WorkBook.Worksheets[$WorksheetName].Activate()
    
                        # SET REFERENCE TO WORKSHEET OBJECT
                        $WorkSheet = $WorkBook.Worksheets[$WorksheetName]
    
                        #region NUMBER OF ROWS IN WORKSHEET
                        $RowCount = 0
                        $RowCount = $WorkSheet.UsedRange.Rows.Count - 1
                        $Notice = "Number of original rows: '" + $RowCount + "'"
                        # [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
                        Write-Output ""
                        Write-Output $Notice
                        #endregion NUMBER OF ROWS IN WORKSHEET
    
                        #region REMOVE UNIDENTIFIED COLUMNS
                        if (-not ($RetainedHeadersArray.Header[0] -eq "*")) {
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
                        }
                        #endregion REMOVE UNIDENTIFIED COLUMNS
                    }
                    else {
                        # ADD NEW WORKSHEET
                        # $WorkBook.Worksheets.Add()
    
                        # NUMBER OF WORKSHEETS IN WORKBOOK
                        # $intWorksheetCount = 0
                        # $intWorksheetCount = $WorkBook.Worksheets.Count
    
                        # # ADD NEW WORKSHEET TO WORKBOOK
                        # # $WorkBook.Worksheets.Add($WorksheetName)
                        # # $WorkBook.Worksheets.Add([System.Reflection.Missing]::Value, $WorkBook.Worksheets[$intWorksheetCount], 1, $WorkBook.Worksheets[1].Type)
    
                        # # COPY PRIMARY WORKSHEET (TO LAST POSITION)
                        # $WorkBook.Worksheets[1].Copy([System.Reflection.Missing]::Value, $WorkBook.Worksheets[$intWorksheetCount])
    
                        # RENAME WORKSHEET (IN LAST POSITION)
                        # $WorkBook.Worksheets[$intWorksheetCount + 1].Name = $WorksheetName
    
                        # # ACTIVATE WORKSHEET
                        # $WorkBook.Worksheets[$WorksheetName].Activate()
    
                        # # SET REFERENCE TO WORKSHEET OBJECT
                        # $WorkSheet = $WorkBook.Worksheets[$WorksheetName]
                    }
                }
                break
            }

            1 { # ACTION
                $Action = $ActionItem

                switch ($Action) {
                    "FILTER-TEXT" {
                        break
                    }

                    "FILTER-COLOR" {
                        break
                    }

                    "FILTER-CLEAR" {
                        break
                    }

                    "SORT" {
                        break
                    }

                    "COLORIZE" {
                        break
                    }

                    "COPY-PASTE" {
                        break
                    }

                    Default {

                    }
                }
                break
            }

            2 { # FIELD
                $Field = $ActionItem

                switch ($Action) {
                    "FILTER-TEXT" {
                        # FIND HEADER COLUMN INDEX
                        $FilterFieldIndex = $WorkSheet.UsedRange.Columns.Find($Field).Column # Example: "Country = 2nd column"

                        # Add-Content log.txt -Value ""
                        # Add-Content log.txt -Value $Field

                        # Add-Content log.txt -Value "Start ForEach"
                        # $Timestamp = $(Get-Date -Format "MM/dd/yyyy hh:mm:ss tt").ToString()
                        # Add-Content log.txt -Value $Timestamp

                        # POPULATE COLUMN TEXT ARRAY WITH COLUMN CONTENTS
                        $ColumnTextArray = $WorkSheet.UsedRange.Rows.Columns[$FilterFieldIndex].Cells | ForEach-Object { $_.Text }

                        # Add-Content log.txt -Value "End ForEach"
                        # $Timestamp = $(Get-Date -Format "MM/dd/yyyy hh:mm:ss tt").ToString()
                        # Add-Content log.txt -Value $Timestamp

                        # REMOVE HEADER COLUMN FROM ARRAY
                        $ColumnTextArray = $ColumnTextArray | Select-Object -skip 1

                        # DE-DUPE AND SORT
                        $ColumnTextArray = $ColumnTextArray | Sort-Object -Unique # Example: 8 rows
                        break
                    }

                    "FILTER-COLOR" {
                        break
                    }

                    "FILTER-CLEAR" {
                        break
                    }

                    "SORT" {
                        break
                    }

                    "COLORIZE" {
                        if ($Field -eq "*") { # Note: "*" represents all columns.
                            $FilterFieldIndex = -1
                        }
                        break
                    }

                    "COPY-PASTE" {
                        break
                    }

                    Default {

                    }
                }
                break
            }

            3 { # CRITERIA
                $Criteria = $ActionItem

                # REMOVE SURROUNDING QUOTATION MARKS
                $Criteria = $Criteria.Trim('"')

                # REMOVE LEADING BRACKET
                $Criteria = $Criteria.TrimStart('[')

                # REMOVE TRAILING BRACKET
                $Criteria = $Criteria.TrimEnd(']')

                $CriteriaArray = @()
                $CriteriaArray = $Criteria.Split('|')

                switch ($Action) {
                    "FILTER-TEXT" {
                        $MatchingArrayItems = @()
                        foreach ($CriteriaItem in $CriteriaArray) {
                            # SEARCH FIELD ARRAY
                            $MatchingArrayItems = ReturnMatchingArrayItems $ColumnTextArray $CriteriaItem # Example: Filter = "Can*; Result = "Canada"

                            # ADD TO AUTOFILTER CRITERIA
                            $FilterCriteriaArray += $MatchingArrayItems
                        }
                    
                        # APPLY AUTOFILTER
                        $WorkSheet.UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints number of records found.

                        # SELECT ALL FILTERED CELLS
                        # Note: "SpecialCells" returns a non-contiguous range. You must loop through each area.
                        # $FilteredCells = $null
                        $FilteredCells = $WorkSheet.AutoFilter.Range.SpecialCells($xlCellTypeVisible)
                        $FilteredCells.Select() # Prints "True" if successful.

                        #region NUMBER OF FILTERED WORKSHEET ROWS
                        $FilteredRowCount = -1
                        foreach ($FilteredRow in $FilteredCells.Rows) {
                            $FilteredRowCount++
                        }
                        $Notice = "Number of filtered rows: '" + $FilteredRowCount + "'"
                        # [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
                        Write-Output ""
                        Write-Output $Notice
                        #endregion NUMBER OF FILTERED WORKSHEET ROWS
                        break
                    }

                    "FILTER-COLOR" {
                        break
                    }

                    "FILTER-CLEAR" {
                        break
                    }

                    "SORT" {
                        break
                    }

                    "COLORIZE" {
                        # USE FILL COLOR FOR FILTERED CELLS
                        if ($FilterFieldIndex -eq -1) { # Note: "-1" represents all columns.
                            foreach ($FilteredRow in $FilteredCells.Rows) {
                                if ($FilteredRow.Row -gt 1) { # Skip header row
                                    # $FilteredRow.Interior.ColorIndex = $yellow # Yellow = 6
                                    $FilteredRow.Interior.ColorIndex = $ColorTable[$Criteria]
                                }
                            }
                        }
                        else {
                            # COLORIZE A PARTICULAR COLUMN OR CELL
                        }
                        break
                    }

                    "COPY-PASTE" {
                        # ADD NEW WORKSHEET
                        $WorkBook.Worksheets.Add()

                        # NUMBER OF WORKSHEETS IN WORKBOOK
                        $intWorksheetCount = 0
                        $intWorksheetCount = $WorkBook.Worksheets.Count

                        $DestinationWorksheetName = $CriteriaArray[0].Split('=')[1] # Example: "Test1"
                        $CopyHeaders = $CriteriaArray[1].Split('=')[1] # Example: "YES/NO"
        
                        # RENAME WORKSHEET (IN LAST POSITION)
                        $WorkBook.Worksheets[$intWorksheetCount + 1].Name = $DestinationWorksheetName

                        # COPY/PASTE
                        # $Source = $WorkSheet
                        # $Destination = $WorkBook.Worksheets[$WorksheetName].Range("A1")
                        # $Source.CopyTo($Destination, ExcelCopyRangeOptions.All)

                        # ACTIVATE WORKSHEET
                        # $WorkBook.Worksheets[$WorksheetName].Activate()

                        # SET REFERENCE TO WORKSHEET OBJECT
                        # $WorkSheet = $WorkBook.Worksheets[$WorksheetName]
                        break
                    }

                    Default {

                    }
                }
                break
            }

            Default {
            }
        }
    }
}
#endregion LOOP THROUGH ACTION ITEMS


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

# RELEASE RESOURCES
While([System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) -ge 0){}
while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) -ge 0){}
