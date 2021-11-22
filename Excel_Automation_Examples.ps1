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


# CELL TYPE CONSTANTS
$xlCellTypeLastCell = 11
$xlCellTypeVisible = 12


# WORKSHEET TYPES
$xlChart = -4109
$xlDialogSheet = -4116
$xlExcel4IntlMacroSheet = 4
$xlExcel4MacroSheet = 3
$xlWorksheet = -4167


# PASTE TYPES
$xlPasteAll = -4104                         # Everything will be pasted.
$xlPasteAllExceptBorders = 7                # Everything except borders will be pasted.
$xlPasteAllMergingConditionalFormats = 14   # Everything will be pasted and conditional formats will be merged.
$xlPasteAllUsingSourceTheme = 13            # Everything will be pasted using the source theme.
$xlPasteColumnWidths = 8                    # Copied column width is pasted.
$xlPasteComments = -4144                    # Comments are pasted.
$xlPasteFormats = -4122                     # Copied source format is pasted.
$xlPasteFormulas = -4123                    # Formulas are pasted.
$xlPasteFormulasAndNumberFormats = 11       # Formulas and Number formats are pasted.
$xlPasteValidation = 6                      # Validations are pasted.
$xlPasteValues = -4163                      # Values are pasted.
$xlPasteValuesAndNumberFormats = 12         # Values and Number formats are pasted.


# PASTE SPECIAL OPERATIONS
$xlPasteSpecialOperationAdd = 2             # Copied data will be added with the value in the destination cell.
$xlPasteSpecialOperationDivide = 5          # Copied data will be divided with the value in the destination cell.
$xlPasteSpecialOperationMultiply = 4        # Copied data will be multiplied with the value in the destination cell.
$xlPasteSpecialOperationNone = -4142        # No calculation will be done in the paste operation.
$xlPasteSpecialOperationSubtract = 3        # Copied data will be subtracted with the value in the destination cell.


# COLORS
$xlNone = -4142


# CELL COLOR INDEX
# $ColorIndexTable = @(
#     No Fill = 0;
#     Black = 1;
#     White = 2;
#     Red = 3;
#     Lime = 4;
#     Blue = 5;
#     Yellow = 6;
#     Fuchsia = 7;
#     Aqua = 8;
#     Maroon = 9;
#     Green = 10;
#     Navy = 11;
#     Olive = 12;
#     Purple = 13;
#     Teal = 14;
#     Silver = 15;
#     Gray = 16;
# )


# CELL COLORS
# $ColorTable = @{
#     Black = "255,0,0,0";
#     White = "255,255,255,255";
#     Red = "255,255,0,0";
#     Lime = "255,0,255,0";
#     Blue = "255,0,0,255";
#     Yellow = "255,255,255,0";
#     Fuchsia = "255,255,0,255";
#     Aqua = "255,0,255,255";
#     Maroon = "255,128,0,0";
#     Green = "255,0,128,0";
#     Navy = "255,0,0,128";
#     Olive = "255,128,128,0";
#     Purple = "255,128,0,128";
#     Teal = "255,0,128,128";
#     Silver = "255,192,192,192";
#     Gray = "255,128,128,128";
#     LimeGreen = "255,50,205,50";
#     PeachPuff = "255,255,218,185"
# }


#region FUNCTIONS
function ReturnMatchingArrayItems {
    # param ($refArray, $refSearchString)
    param ([string[]]$refArray, [string]$refSearchString)

    $refMatchingItemsArray = @()
    
    foreach ($refTextItem in $refArray) {
        if ($refSearchString -eq '""') { # Is search string empty?
        # if ($refSearchString -eq '""""') { # Is search string empty? (Note: This works for both Excel and Notepad.)
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
#endregion FUNCTIONS


#region CLASSES
class SETTINGS {
    [string]$Password
    [string]$FilePath
    [string]$HeadersPath
    [string]$ActionsPath

    # CONSTRUCTOR
    SETTINGS([string]$refFilePathSettings) {
        [System.Array]$SettingsObject = Import-Csv -Path $refFilePathSettings -Delimiter ","
        $this.Password = $SettingsObject.PASSWORD
        $this.FilePath = $SettingsObject.DIRECTORY + $SettingsObject.FILENAME
        $this.HeadersPath = $SettingsObject.DIRECTORY + $SettingsObject.HEADERS
        $this.ActionsPath = $SettingsObject.DIRECTORY + $SettingsObject.ACTIONS
    }
}


class ARGB_Color_Object {
    [string] hidden $color
    [int] hidden $alpha
    [int] hidden $red
    [int] hidden $green
    [int] hidden $blue

    # CELL COLORS
    [System.Collections.Hashtable] hidden $ColorTable = @{
        Black = "255,0,0,0";
        White = "255,255,255,255";
        Red = "255,255,0,0";
        Lime = "255,0,255,0";
        Blue = "255,0,0,255";
        Yellow = "255,255,255,0";
        Fuchsia = "255,255,0,255";
        Aqua = "255,0,255,255";
        Maroon = "255,128,0,0";
        Green = "255,0,128,0";
        Navy = "255,0,0,128";
        Olive = "255,128,128,0";
        Purple = "255,128,0,128";
        Teal = "255,0,128,128";
        Silver = "255,192,192,192";
        Gray = "255,128,128,128";
        LimeGreen = "255,50,205,50";
        PeachPuff = "255,255,218,185"
    }

    # CONSTRUCTOR
    # ARGB_Color_Object([string]$Color) {
    #     $this.SetColor($Color)
    # }

    [void]SetColor([string]$Color) {
        $this.color = $Color
        $this.alpha = $this.ColorTable[$Color].Split(',')[0]
        $this.red = $this.ColorTable[$Color].Split(',')[1]
        $this.green = $this.ColorTable[$Color].Split(',')[2]
        $this.blue = $this.ColorTable[$Color].Split(',')[3]
    }

    [string]GetColor() {
        # $refColor = $this.alpha + ',' + $this.red + ',' + $this.green + ',' + $this.blue
        $refColor = $this.color
        return $refColor
    }

    [System.Drawing.Color]GetRgbColorObject() {
        $refRgbColorObject = [System.Drawing.Color]::FromArgb($this.red, $this.green, $this.blue)
        return $refRgbColorObject
    }
    
    [System.Drawing.Color]GetArgbColorObject() {
        $refArgbColorObject = [System.Drawing.Color]::FromArgb($this.alpha, $this.red, $this.green, $this.blue)
        return $refArgbColorObject
    }
}
#endregion CLASSES


# WORKSHEET NAME
$WorksheetName = ""

# FILTERED CELL ARRAY
$FilteredCells = $null

# DATE INFO
# $Date = Get-Date -Format "yyyy-MM-dd"

# SETTINGS FILEPATH
# $FileDirectory = "C:\Projects\PowerShell\Excel_Automation\"
# $Filename = "Financial Sample"
# $ModifiedFilename = "Financial Sample"
# $ModifiedFilename = "Modified"
# $FilePath = $FileDirectory + $Filename + "." + "xlsx"
# $FilePathBackup = $FileDirectory + $Filename + ".bak" + "." + "xlsx"
# $FilePathModified = $FileDirectory + $ModifiedFilename + " " + $Date + "." + "xlsx"
# $FilePathRetainedHeaders = $FileDirectory + $Filename + " - Headers" + "." + "csv"
# $FilePathActionItems = $FileDirectory + $Filename + " - Filters" + "." + "csv"
# $FilePathActionItems = $FileDirectory + $Filename + " - Actions" + "." + "csv"
$FilePathSettings = "Settings.csv"

# APP SETTINGS
# $Settings = Import-Csv -Path $FilePathSettings -Delimiter ","
$objSettings = [SETTINGS]::new($FilePathSettings)

# EXCEL WORKBOOK PASSWORD
# $WorkbookPassword = ""
$WorkbookPassword = $objSettings.Password

# COLUMNS IDENTIFIED FOR RETENTION
# $RetainedHeadersArray = @("Header", "Segment", "Country", "Product", "Date", "Month Number", "Month Name", "Year")
# $RetainedHeadersArray = Import-Csv -Path $FilePathRetainedHeaders -Delimiter ","
$RetainedHeadersArray = Import-Csv -Path $objSettings.HeadersPath -Delimiter ","

# ACTION ITEMS FILE
# $ActionsArray = Get-Content -Path $FilePathActionItems
# $ActionsArray = Import-Csv -Path $FilePathActionItems -Delimiter ","
$ActionsArray = Import-Csv -Path $objSettings.ActionsPath -Delimiter ","

# WORKING EXCEL FILE
$FilePath = $objSettings.FilePath

# BACKUP FILE(S)
# Copy-Item $FilePath -Destination $FilePathBackup -Force
# Copy-Item $FilePath -Destination $FilePathModified -Force

# Note: The following is necessary for such things as "MessageBox".
# Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $true

# OPEN THE EXCEL FILE
$Workbook = $objExcel.Workbooks.Open($FilePath)
# $Workbook = $objExcel.Workbooks.Open($FilePathModified)


#region BACKUP WORKBOOK (WITH PASSWORD)
if (-not ([string]::IsNullOrWhiteSpace($WorkbookPassword))) {
    $objExcel.DisplayAlerts = $false; # Note: "$false" = Do not prompt for confirmation to "over-write" backup file.
    $Workbook.SaveAs($FilePath, [Type]::Missing, $WorkbookPassword)
    # $Workbook.SaveAs($FilePathModified, [Type]::Missing, $WorkbookPassword)
    $objExcel.DisplayAlerts = $true; # Note: "$true" = Prompt for confirmation to "over-write" backup file.
}
#endregion BACKUP WORKBOOK (WITH PASSWORD)


#region REMOVE UNIDENTIFIED COLUMNS
# if (-not ($RetainedHeadersArray.HEADER[0] -eq "*")) {
if (-not ($RetainedHeadersArray[0].HEADER -eq "*")) {
    $ColumnIndex = 1
    # $FieldCount = $Workbook.Worksheets[1].UsedRange.Rows(1).Cells.Count
    $FieldCount = $Workbook.Worksheets[1].UsedRange.Rows(1).Columns.Count

    while ($ColumnIndex -le $FieldCount) {
        # $FieldCell = $Workbook.Worksheets[1].UsedRange.Rows(1).Cells[$ColumnIndex]
        $FieldCell = $Workbook.Worksheets[1].UsedRange.Rows(1).Columns[$ColumnIndex].Cells
        $ColumnText = $FieldCell.Text

        # if ($RetainedHeadersArray -match $ColumnText) {
        if ($RetainedHeadersArray.HEADER -match $ColumnText) {
            # $ResponseText = "Field Header: '" + $ColumnText + "' should be retained."
            $ColumnIndex++
        }
        else {
            # $ResponseText = "Field Header: '" + $ColumnText + "' should be DELETED."
            $Workbook.Worksheets[1].Columns[$ColumnIndex].EntireColumn.Delete() # Prints "True" if successful.
            # $FieldCount = $Workbook.Worksheets[1].UsedRange.Rows(1).Cells.Count
            $FieldCount = $Workbook.Worksheets[1].UsedRange.Rows(1).Columns.Count
        }
        # Write-Output $ResponseText
    }
}
#endregion REMOVE UNIDENTIFIED COLUMNS


#region NUMBER OF ROWS IN WORKSHEET
$RowCount = 0
$RowCount = $Workbook.Worksheets[1].UsedRange.Rows.Count - 1
$Notice = "Number of original rows: '" + $RowCount + "'"
# [System.Windows.MessageBox]::Show($Notice) # Prints results of selection, based upon MessageBox options.
Write-Output ""
Write-Output $Notice
#endregion NUMBER OF ROWS IN WORKSHEET


#region LOOP THROUGH ACTION ITEMS
# READ EACH RECORD INTO ARRAY
# for ($intRecordCounter = 0; $intRecordCounter -lt $ActionsArray.Count; $intRecordCounter++) {
for ($intRecordCounter = 0; $intRecordCounter -lt $($ActionsArray | Measure-Object).Count; $intRecordCounter++) {
    # CLEAR AUTOFILTER CRITERIA
    [string[]]$FilterCriteriaArray = @()

    # CLEAR ACTION
    $Action = ""

    # CLEAR REFINED COLUMN CONTENTS
    $ColumnTextArray = @()

    # CLEAR FILTER FIELD INDEX
    $FilterFieldIndex = 0

    # $WorksheetName = $ActionsArray.WORKSHEET[$intRecordCounter]
    $WorksheetName = $ActionsArray[$intRecordCounter].WORKSHEET

    # $Action = $ActionsArray.ACTION[$intRecordCounter]
    $Action = $ActionsArray[$intRecordCounter].ACTION

    # $Field = $ActionsArray.FIELD[$intRecordCounter]
    $Field = $ActionsArray[$intRecordCounter].FIELD

    # $Criteria = $ActionsArray.CRITERIA[$intRecordCounter]
    $Criteria = $ActionsArray[$intRecordCounter].CRITERIA

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
            # FIND HEADER COLUMN INDEX
            # $FilterFieldIndex = $Workbook.Worksheets[$WorksheetName].UsedRange.Columns.Find($Field).Column # Example: "Country = 2nd column"
            $FilterFieldIndex = $Workbook.Worksheets[$WorksheetName].UsedRange.Rows.Columns.Find($Field).Column # Example: "Country = 2nd column"

            # Add-Content log.txt -Value ""
            # Add-Content log.txt -Value $Field

            # Add-Content log.txt -Value "Start ForEach"
            # $Timestamp = $(Get-Date -Format "MM/dd/yyyy hh:mm:ss tt").ToString()
            # Add-Content log.txt -Value $Timestamp

            # POPULATE COLUMN TEXT ARRAY WITH COLUMN CONTENTS
            $ColumnTextArray = $Workbook.Worksheets[$WorksheetName].UsedRange.Rows.Columns[$FilterFieldIndex].Cells | ForEach-Object { $_.Text }

            # Add-Content log.txt -Value "End ForEach"
            # $Timestamp = $(Get-Date -Format "MM/dd/yyyy hh:mm:ss tt").ToString()
            # Add-Content log.txt -Value $Timestamp

            # REMOVE HEADER COLUMN FROM ARRAY
            $ColumnTextArray = $ColumnTextArray | Select-Object -skip 1

            # DE-DUPE AND SORT
            $ColumnTextArray = $ColumnTextArray | Sort-Object -Unique # Example: 8 rows
            
            $MatchingArrayItems = @()
            foreach ($CriteriaItem in $CriteriaArray) {
                # SEARCH FIELD ARRAY
                $MatchingArrayItems = ReturnMatchingArrayItems $ColumnTextArray $CriteriaItem # Example: Filter = "Can*; Result = "Canada"

                # ADD TO AUTOFILTER CRITERIA
                $FilterCriteriaArray += $MatchingArrayItems
            }
        
            # APPLY AUTOFILTER
            $Workbook.Worksheets[$WorksheetName].UsedRange.AutoFilter($FilterFieldIndex, $FilterCriteriaArray, $xlFilterValues) # Prints index of column where filter was applied.

            # SELECT ALL FILTERED CELLS
            # Note: "SpecialCells" returns a non-contiguous range. You must loop through each area.
            $FilteredCells = $null
            $FilteredCells = $Workbook.Worksheets[$WorksheetName].AutoFilter.Range.SpecialCells($xlCellTypeVisible)
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
            # FIND HEADER COLUMN INDEX
            # $FilterFieldIndex = $Workbook.Worksheets[$WorksheetName].UsedRange.Columns.Find($Field).Column # Example: "Country = 2nd column"
            $FilterFieldIndex = $Workbook.Worksheets[$WorksheetName].UsedRange.Rows.Columns.Find($Field).Column # Example: "Country = 2nd column"
        
            # # RGB COLOR FILTER
            # $RgbColor = [System.Drawing.Color]::FromArgb($CriteriaArray[0], $CriteriaArray[1], $CriteriaArray[2])
            $objRgbColor = [ARGB_Color_Object]::new()
            $objRgbColor.SetColor($Criteria)
            $RgbColor = $objRgbColor.GetRgbColorObject()

            # APPLY AUTOFILTER
            $Workbook.Worksheets[$WorksheetName].UsedRange.AutoFilter($FilterFieldIndex, $RgbColor, $xlFilterCellColor) # Prints index of column where filter was applied.

            # SELECT ALL FILTERED CELLS
            # Note: "SpecialCells" returns a non-contiguous range. You must loop through each area.
            $FilteredCells = $null
            $FilteredCells = $Workbook.Worksheets[$WorksheetName].AutoFilter.Range.SpecialCells($xlCellTypeVisible)
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

        "FILTER-CLEAR" {
            if ($Field -eq "*") { # Note: "*" represents all columns.
                # CLEAR AUTOFILTER
                $Workbook.Worksheets[$WorksheetName].UsedRange.AutoFilter() # Prints True if filter was successful.
            }
            else {
                # FIND HEADER COLUMN INDEX
                $FilterFieldIndex = $Workbook.Worksheets[$WorksheetName].UsedRange.Columns.Find($Field).Column # Example: "Country = 2nd column"
        
                # CLEAR AUTOFILTER
                $Workbook.Worksheets[$WorksheetName].UsedRange.AutoFilter($FilterFieldIndex) # Prints True if filter was successful.
            }
            break
        }

        "SORT" {
            break
        }

        "COLOR-FILL" {
            # USE FILL COLOR FOR FILTERED CELLS
            if (-not ($Criteria -eq "No Fill")) {
                $objRgbColor = [ARGB_Color_Object]::new()
                $objRgbColor.SetColor($Criteria)
                $RgbColor = $objRgbColor.GetRgbColorObject()
            }

            if ($Field -eq "*") { # Note: "*" represents all columns.
                foreach ($FilteredRow in $FilteredCells.Rows) {
                    if ($FilteredRow.Row -gt 1) { # Skip header row
                        if ($Criteria -eq "No Fill") {
                            $FilteredRow.Interior.ColorIndex = 0
                            # $FilteredRow.Interior.ColorIndex = $xlNone
                        }
                        else {
                            $FilteredRow.Interior.Color = $RgbColor
                        }
                    }
                }
            }
            else {
                # COLOR A PARTICULAR COLUMN OR CELL
            }
            break
        }

        "ADD-WORKSHEET" {
            # NUMBER OF WORKSHEETS IN WORKBOOK
            $intWorksheetCount = 0
            $intWorksheetCount = $Workbook.Worksheets.Count

            # ADD NEW WORKSHEET TO WORKBOOK
            $NewWorksheet = $null
            # $Workbook.Worksheets.Add([Before Sheet], [After Sheet], [Number of Sheets to be Added], [Sheet Type])
            $NewWorksheet = $Workbook.Worksheets.Add([Type]::Missing, $Workbook.Worksheets[$intWorksheetCount], 1, $xlWorksheet)

            # RENAME NEW WORKSHEET
            # $Workbook.Worksheets[$intWorksheetCount + 1].Name = $NewWorksheetName
            $NewWorksheet.Name = $WorksheetName

            # ACTIVATE MAIN WORKSHEET
            $Workbook.Worksheets[1].Activate()
            break
        }

        "COPY-PASTE" {
            if ($Field -eq "NULL") {

            }
            
            # COPY FILTERED CELLS
            $Workbook.Worksheets[$WorksheetName].AutoFilter.Range.SpecialCells($xlCellTypeVisible).Copy()

            $NewWorksheetName = ""
            $CopyHeaders = $false

            $NewWorksheetName = $CriteriaArray[0].Split('=')[1] # Example: "Test1"
            $CopyHeaders = $CriteriaArray[1].Split('=')[1] # Example: "TRUE/FALSE"

            # PASTE CELLS
            # $Workbook.Worksheets[$NewWorksheet.Name].Range("A1").PasteSpecial($xlPasteValues, $xlPasteSpecialOperationNone, $false, $false)
            $Workbook.Worksheets[$NewWorksheetName].Range("A1").PasteSpecial($xlPasteAll, $xlPasteSpecialOperationNone, $false, $false)

            # GET LAST CELL IN PASTE AREA
            # $LastCell = $objRange.SpecialCells($xlCellTypeLastCell)
            $LastCell = $Workbook.Worksheets[$NewWorksheetName].UsedRange.SpecialCells($xlCellTypeLastCell)

            # ACTIVATE DESTINATION WORKSHEET
            $Workbook.Worksheets[$NewWorksheetName].Activate()

            # MOVE TO LAST ROW, FIRST COLUMN ON DESTINATION WORKSHEET (FOR POSSIBLE FUTURE COPY/PASTE OPERATIONS)
            $Workbook.Worksheets[$NewWorksheetName].UsedRange.Rows[$LastCell.Row + 1].Columns[1].Cells.Select()

            # ACTIVATE MAIN WORKSHEET
            $Workbook.Worksheets[1].Activate()
            break
        }

        Default {

        }
    }
}
#endregion LOOP THROUGH ACTION ITEMS

# TURN OFF CLIPBOARD WARNING MESSAGE
$Workbook.Application.CutCopyMode = $false

# CLOSE WORKBOOK
$Workbook.Close($true) # Note: "$true" = Save file changes. "$false" = Do not save file changes.

# EXIT EXCEL
$objExcel.Quit()

# RELEASE RESOURCES
While([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) -ge 0){}
while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) -ge 0){}
