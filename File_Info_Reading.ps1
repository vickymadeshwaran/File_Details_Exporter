#************************************************************
# Project Name: File Details Extractor
# Build Environment: Windows PowerShell ISE
# Author: VIGNESH M
# Last modified: 17-11-2024
#*************************************************************


# Prompt the user for input
$sourcePath = Read-Host "Enter the source folder path (e.g., C:\Users\YourName\Documents)"
$extensions = Read-Host "Enter the file extension(s), separate multiple extensions with commas (e.g., .txt,.jpg,.docx)"
$destinationPath = Read-Host "Enter the destination folder path (e.g., C:\Users\YourName\Desktop)"
$excelFileName = Read-Host "Enter the name for the Excel file (without extension, e.g., FileDetails)"

# Check if the source path exists
if (-Not (Test-Path $sourcePath)) {
    Write-Host "The specified source path does not exist. Please check and try again." -ForegroundColor Red
    exit
}

# Check if the destination path exists
if (-Not (Test-Path $destinationPath)) {
    Write-Host "The specified destination path does not exist. Please check and try again." -ForegroundColor Red
    exit
}

# Process the extensions input
$extensionList = $extensions -split ',' | ForEach-Object { $_.Trim() }

# Initialize an empty array to hold file details
$fileDetails = @()

# Loop through each extension and get matching files
foreach ($extension in $extensionList) {
    $files = Get-ChildItem -Path $sourcePath -Filter "*$extension" -File -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        # Extract required details
        $fileName = $file.Name
        $size = "{0:N0}" -f ($file.Length) # Format size as an integer with commas
        $date = $file.LastWriteTime.ToString("dd/MM/yyyy") # Format: DD/MM/YYYY
        $time = $file.LastWriteTime.ToString("HH:mm:ss") # Format: HH:MM:SS
        $ampm = $file.LastWriteTime.ToString("tt") # AM/PM format

        # Add file details as a row
        $fileDetails += [PSCustomObject]@{
            "File Name" = $fileName
            "Size"      = $size
            "Date"      = $date
            "Time"      = $time
            "AM/PM"     = $ampm
        }
    }
}

# Check if any files were found
if ($fileDetails.Count -eq 0) {
    Write-Host "No files matching the specified extension(s) were found in the source path." -ForegroundColor Yellow
    exit
}

# Define the full path for the Excel file
$excelFilePath = Join-Path -Path $destinationPath -ChildPath ("$excelFileName.csv")

# Export the data to a CSV file
try {
    $fileDetails | Export-Csv -Path $excelFilePath -NoTypeInformation -Encoding UTF8
    Write-Host "File details saved successfully to '$excelFilePath'" -ForegroundColor Green
} catch {
    Write-Host "Failed to save the file. Please check the input and try again." -ForegroundColor Red
    Write-Host "Error: $_"
}
