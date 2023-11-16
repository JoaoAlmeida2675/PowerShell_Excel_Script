# Excel Data Validation Automation

Automate the validation of Excel data by checking for blank or empty cells in a specified range of lines and columns. This PowerShell script streamlines the process, making it efficient and easy to identify and manage empty cells in your Excel files.

## Introduction

Managing and validating data in Excel spreadsheets can be a time-consuming task, especially when dealing with large datasets. This PowerShell script simplifies the process by automating the detection of empty cells within a specified range. By leveraging the ImportExcel module, it provides a seamless solution for Excel data validation.

Check INFO.md for more detail information about the code.

## How It Works

### 1. Function Get-ColumnIndexFromLetter:

This function converts a letter or set of letters to uppercase and calculates the corresponding column index. It ensures consistency in identifying columns, where A=1, B=2, etc. By receiving a string `$letter` as a parameter, it iterates over each character, calculates the corresponding index in the alphabet, and returns the column index.

### 2. Open and Load Excel File:

Initiates an instance of Excel, makes it visible (optional), and opens an Excel workbook from the specified path. The `$excel` variable contains the Excel instance, and `$workbook` contains the opened workbook.

### 3. Select Sheet and Specify Columns to Check:

Selects a specific sheet in the workbook and specifies a set of columns to be checked. The `$sheet` variable contains the selected sheet, and `$columnsToCheck` contains the columns to be validated.

### 4. Check Empty Cells:

Iterates over the specified columns and rows to check if the cells are empty. It identifies empty cells, records the column names and messages for empty columns in the `$emptyColumnMessages` variable.

### 5. Create the .txt file with messages for empty columns:

Creates a text file with messages for empty columns. The script specifies the path of the text file and uses `Out-File` to write the messages to the file.

### 6. Close Excel and Release Resources:

Closes the workbook and Excel, releasing resources associated with the Excel objects. It ensures efficient resource management after the validation process.

### 7. Remove COM Object References:

Removes variables storing references to COM objects to further release resources. This step is crucial for preventing resource leaks and maintaining system stability.

## Usage

1. Install the ImportExcel module:
    ```powershell
    Install-Module -Name ImportExcel
    ```

2. Customize the script:
    - Set the Excel file path: `$excelFilePath = "Your\Path\To\Excel\File.xlsx"`
    - Specify the sheet name: `$sheetName = "Your_Sheet_Name"`
    - Define the range of lines and columns to check.

3. Run the script:
    ```powershell
    .\ValidateExcelData.ps1
    ```

Feel free to modify the script according to your specific Excel file and validation requirements.
