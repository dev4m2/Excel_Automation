# Excel_Automation_Examples

## Description
Manipulating Excel files through PowerShell and 
Excel Automation (i.e. COM objects).

STEPS (Identified in csv file)
1. Select an "Active" worksheet.
<br/>

2. Actionable options: 
<br/>
    - <b>FILTER-TEXT</b>
        - Worksheet Name
        - Column Name (Example: "Country")
        - Criteria List (Note: Allows wildcards. Example: "USA", Can\*", "Germ\*")
        <br/><br/>
    - <b>FILTER-COLOR</b>
        - Worksheet Name
        - Column Name (Example: "Product")
        - Cell Color (Example: "Yellow", "Peach", "Blue", "No Fill")
        <br/><br/>
    - <b>FILTER-CLEAR</b>
        - Worksheet Name
        - Column Name (Note: "\*" represents all columns. Example: "Country" OR "\*")
        <br/><br/>
    - <b>SORT</b>
        <br/><br/>
    - <b>COLORIZE</b>
        - Worksheet Name
        - Column Name (Note: "\*" represents all columns. Example: "Date" OR "\*")
        - Selection (Note: Based on prior cell selection, such as through filtering.)
        - Color (Example: "Yellow", "Peach", "Blue", "No Fill")
        <br/><br/>
    - <b>COPY-PASTE</b>
        - Worksheet Name
        - Column Name = NULL
        - Source = Worksheet Name
        - Selection (Note: Based on prior cell selection, such as through filtering.)
        - Destination = Worksheet Name
        - Headers (Include) = "YES/NO"
        <br/><br/>



Python code examples for reading a 'csv' file into a "list of lists".<br/>
The samples present code for referencing such things as: 
1. Count the number of list items (i.e. "rows") from the data object.
2. The ability to search for a string value within the data and return the associated list/row.
3. Return a "search results" index.
4. Return a specific cell value from said list/row.
5. Return a random list of indices from the data object.
<br/><br/>

## Getting Started
### Step 1
Move to the appropriate directory and type the following at the command prompt...<br/>
_**~/Projects/Python/import_csv_examples>**_ **`python3 import_csv_examples.py`**
<br/><br/>

## Reference Material
Download the Financial Sample Excel workbook for Power BI<br/>
https://docs.microsoft.com/en-us/power-bi/create-reports/sample-financial-download

Country Names And Country Codes Reference Lists</br>
https://home.treasury.gov/data/treasury-international-capital-tic-system-home-page/using-tic/country-names-and-country-codes-reference-lists
