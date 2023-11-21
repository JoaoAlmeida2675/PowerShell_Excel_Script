# Function to get the column index from the column letter (A=1, B=2, etc.) - supports columns with multiple letters (e.g., AA, AB)
function Get-ColumnIndexFromLetter {
    param (
        [string]$letter
    )
    $letter = $letter.ToUpper()
    $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $columnIndex = 0

    for ($i = 0; $i -lt $letter.Length; $i++) {
        $character = $letter[$i]
        $characterIndex = $alphabet.IndexOf($character) + 1
        $columnIndex = $columnIndex * 26 + $characterIndex
    }

    return $columnIndex
}

# Open Excel and load the file
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true  # You can set this to $false if you want Excel to run in the background
$workbook = $excel.Workbooks.Open("C:\Users\jalmeida26\OneDrive - DXC Production\Desktop\Projeto ValidacaoDados\WorkFile.xlsx")

if ($workbook -eq $null) {
    Write-Host "Failed to open the workbook."
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    return
}

# Specify the sheet name
$sheetName = "Devices"

# Select the sheet by name
$sheet = $workbook.Sheets.Item($sheetName)

if ($sheet -eq $null) {
    Write-Host "Failed to select the worksheet."
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    return
}

$sheet.Select()

# Specify the columns to check (based on column letters)
$columnsToCheck = "B", "C", "D", "H", "I", "L", "M", "N", "O", "Q", "R", "T", "U", "Y", "Z", "AA", "AD", "AH", "AI", "AJ", "AK", "AM", "AN", "AO"

# Specify the rows to check
$rowsToCheck = 6 # Example: Check rows from 6 to 10 = 6..10

# Specify the row with names
$rowWithNames = 1  # Update with the actual row number containing column names

# Initialize a list to store the names of empty columns
$emptyColumnNames = @()

# Function to get the column name from the column index
function Get-ColumnName {
    param (
        [int]$columnIndex
    )
    $dividend = $columnIndex
    $columnName = ''

    while ($dividend -gt 0) {
        $modulo = ($dividend - 1) % 26
        $columnName = [char]([int][char]'A' + $modulo) + $columnName
        $dividend = [int](($dividend - $modulo) / 26)
    }

    return $columnName
}

# Initialize a list to store the messages of empty columns
$emptyColumnMessages = @()

# Loop through columns and rows and get the column name at the specified row
foreach ($columnLetter in $columnsToCheck) {
    $columnIndex = Get-ColumnIndexFromLetter $columnLetter

    if ($columnIndex -eq 0) {
        Write-Host "Invalid column letter: $columnLetter"
        continue
    }

    # Activate the worksheet
    $sheet.Activate()

    # Get the cell at the specified row and column
    $cell = $sheet.Cells.Item($rowWithNames, $columnIndex)

    # Check for null values when accessing the cell
    if ($cell -eq $null) {
        $columnName = "N/A"  # or any default value
    } else {
        $columnName = $cell.Value2
    }

    # Loop through rows and check for missing data
    foreach ($row in $rowsToCheck) {
        $cellValue = $sheet.Cells.Item($row, $columnIndex).Value2

        if ($cellValue -eq $null) {
            $message = "Row: $row - Missing Data in Column: $columnName"
            Write-Host $message
            $emptyColumnMessages += $message
        }
    }
}

# Create the .txt file with messages of empty columns
$fileName = "$env:USERPROFILE\OneDrive - DXC Production\Desktop\MissingDataExcel.txt"
$emptyColumnMessages | Out-File -FilePath $fileName

# Close Excel
$workbook.Close()
$excel.Quit()

# Release resources
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# Remove references to COM objects
Remove-Variable sheet
Remove-Variable workbook
Remove-Variable excel
