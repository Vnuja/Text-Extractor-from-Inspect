# PowerShell HTML Text Extractor Script
# This script extracts text between span tags from HTML elements and saves to text files

function Select-FileDialog {
    param([string]$Title, [string]$Filter = "CSV Files (*.csv)|*.csv")
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $Title
    $openFileDialog.Filter = $Filter
    $openFileDialog.ShowDialog() | Out-Null
    return $openFileDialog.FileName
}

function Select-FolderDialog {
    param([string]$Description = "Select folder to save text files")
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.ShowDialog() | Out-Null
    return $folderBrowser.SelectedPath
}

function Extract-SpanContent {
    param ([string]$htmlContent)
    
    # Pattern to match text between span tags
    $pattern = '<span[^>]*>(.*?)<\/span>'
    $matches = [regex]::Matches($htmlContent, $pattern)
    
    # Join all span contents
    $result = $matches | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -ne "" }
    return ($result -join "`n")
}

# Get the CSV file path
Write-Host "Please select the CSV file containing text file names and HTML elements" -ForegroundColor Cyan
$csvPath = Select-FileDialog -Title "Select CSV file"

if (-not $csvPath -or -not (Test-Path $csvPath)) {
    Write-Host "No file selected or file does not exist. Exiting script." -ForegroundColor Red
    exit
}

# Get the destination folder path
Write-Host "Please select the destination folder where text files will be saved" -ForegroundColor Cyan
$destinationRoot = Select-FolderDialog -Description "Select destination folder for text files"

if (-not $destinationRoot) {
    Write-Host "No destination folder selected. Exiting script." -ForegroundColor Red
    exit
}

# Import CSV file
# Assuming first column is "FileName" and second column is "HtmlElement"
try {
    $csvData = Import-Csv $csvPath
    
    # Check column names and adapt if needed
    $columns = $csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    $fileNameColumn = $columns[0] # First column for file name
    $htmlColumn = $columns[1]     # Second column for HTML content
    
    Write-Host "Using '$fileNameColumn' column for file names and '$htmlColumn' column for HTML content" -ForegroundColor Yellow
} catch {
    Write-Host "Error reading CSV file: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Create a progress counter
$totalItems = $csvData.Count
$processedItems = 0
$successItems = 0
$failedItems = 0

Write-Host "Found $totalItems items to process" -ForegroundColor Green

# Process each row in the CSV
foreach ($row in $csvData) {
    $processedItems++
    
    # Get file name and HTML content
    $fileName = $row.$fileNameColumn
    $htmlContent = $row.$htmlColumn
    
    # Skip if file name or HTML content is empty
    if (-not $fileName -or $fileName -eq "" -or -not $htmlContent -or $htmlContent -eq "") {
        Write-Host "Skipping row $processedItems: Empty file name or HTML content" -ForegroundColor Yellow
        $failedItems++
        continue
    }
    
    # Make sure file name has .txt extension
    if (-not $fileName.EndsWith(".txt")) {
        $fileName = "$fileName.txt"
    }
    
    try {
        # Extract span content
        $extractedText = Extract-SpanContent -htmlContent $htmlContent
        
        # Skip if no text was extracted
        if (-not $extractedText -or $extractedText -eq "") {
            Write-Host "Skipping $fileName: No span content found" -ForegroundColor Yellow
            $failedItems++
            continue
        }
        
        # Create folder with the same name as the text file (without extension)
        $folderName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $folderPath = Join-Path -Path $destinationRoot -ChildPath $folderName
        
        if (-not (Test-Path $folderPath)) {
            New-Item -Path $folderPath -ItemType Directory | Out-Null
            Write-Host "Created folder: $folderName" -ForegroundColor Yellow
        }
        
        # Set the file path
        $filePath = Join-Path -Path $folderPath -ChildPath $fileName
        
        # Write the extracted text to the file
        $extractedText | Out-File -FilePath $filePath -Encoding utf8
        $successItems++
        
        # Display progress
        $progressPercent = [math]::Round($processedItems / $totalItems * 100, 2)
        Write-Host "Processed ($progressPercent%): Created $fileName with extracted span text" -ForegroundColor Green
        
    } catch {
        Write-Host "Error processing $fileName: $($_.Exception.Message)" -ForegroundColor Red
        $failedItems++
    }
}

# Display summary
Write-Host "`nExtraction Summary:" -ForegroundColor Cyan
Write-Host "Total items processed: $processedItems" -ForegroundColor White
Write-Host "Successfully extracted: $successItems" -ForegroundColor Green
Write-Host "Failed or skipped: $failedItems" -ForegroundColor Yellow
Write-Host "`nText extraction completed. Files saved to: $destinationRoot" -ForegroundColor Green# PowerShell HTML Text Extractor Script
# This script extracts text between span tags from HTML elements and saves to text files

function Select-FileDialog {
    param([string]$Title, [string]$Filter = "CSV Files (*.csv)|*.csv")
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $Title
    $openFileDialog.Filter = $Filter
    $openFileDialog.ShowDialog() | Out-Null
    return $openFileDialog.FileName
}

function Select-FolderDialog {
    param([string]$Description = "Select folder to save text files")
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.ShowDialog() | Out-Null
    return $folderBrowser.SelectedPath
}

function Extract-SpanContent {
    param ([string]$htmlContent)
    
    # Pattern to match text between span tags
    $pattern = '<span[^>]*>(.*?)<\/span>'
    $matches = [regex]::Matches($htmlContent, $pattern)
    
    # Join all span contents
    $result = $matches | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -ne "" }
    return ($result -join "`n")
}

# Get the CSV file path
Write-Host "Please select the CSV file containing text file names and HTML elements" -ForegroundColor Cyan
$csvPath = Select-FileDialog -Title "Select CSV file"

if (-not $csvPath -or -not (Test-Path $csvPath)) {
    Write-Host "No file selected or file does not exist. Exiting script." -ForegroundColor Red
    exit
}

# Get the destination folder path
Write-Host "Please select the destination folder where text files will be saved" -ForegroundColor Cyan
$destinationRoot = Select-FolderDialog -Description "Select destination folder for text files"

if (-not $destinationRoot) {
    Write-Host "No destination folder selected. Exiting script." -ForegroundColor Red
    exit
}

# Import CSV file
# Assuming first column is "FileName" and second column is "HtmlElement"
try {
    $csvData = Import-Csv $csvPath
    
    # Check column names and adapt if needed
    $columns = $csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    $fileNameColumn = $columns[0] # First column for file name
    $htmlColumn = $columns[1]     # Second column for HTML content
    
    Write-Host "Using '$fileNameColumn' column for file names and '$htmlColumn' column for HTML content" -ForegroundColor Yellow
} catch {
    Write-Host "Error reading CSV file: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Create a progress counter
$totalItems = $csvData.Count
$processedItems = 0
$successItems = 0
$failedItems = 0

Write-Host "Found $totalItems items to process" -ForegroundColor Green

# Process each row in the CSV
foreach ($row in $csvData) {
    $processedItems++
    
    # Get file name and HTML content
    $fileName = $row.$fileNameColumn
    $htmlContent = $row.$htmlColumn
    
    # Skip if file name or HTML content is empty
    if (-not $fileName -or $fileName -eq "" -or -not $htmlContent -or $htmlContent -eq "") {
        Write-Host "Skipping row $processedItems: Empty file name or HTML content" -ForegroundColor Yellow
        $failedItems++
        continue
    }
    
    # Make sure file name has .txt extension
    if (-not $fileName.EndsWith(".txt")) {
        $fileName = "$fileName.txt"
    }
    
    try {
        # Extract span content
        $extractedText = Extract-SpanContent -htmlContent $htmlContent
        
        # Skip if no text was extracted
        if (-not $extractedText -or $extractedText -eq "") {
            Write-Host "Skipping $fileName: No span content found" -ForegroundColor Yellow
            $failedItems++
            continue
        }
        
        # Create folder with the same name as the text file (without extension)
        $folderName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $folderPath = Join-Path -Path $destinationRoot -ChildPath $folderName
        
        if (-not (Test-Path $folderPath)) {
            New-Item -Path $folderPath -ItemType Directory | Out-Null
            Write-Host "Created folder: $folderName" -ForegroundColor Yellow
        }
        
        # Set the file path
        $filePath = Join-Path -Path $folderPath -ChildPath $fileName
        
        # Write the extracted text to the file
        $extractedText | Out-File -FilePath $filePath -Encoding utf8
        $successItems++
        
        # Display progress
        $progressPercent = [math]::Round($processedItems / $totalItems * 100, 2)
        Write-Host "Processed ($progressPercent%): Created $fileName with extracted span text" -ForegroundColor Green
        
    } catch {
        Write-Host "Error processing $fileName: $($_.Exception.Message)" -ForegroundColor Red
        $failedItems++
    }
}

# Display summary
Write-Host "`nExtraction Summary:" -ForegroundColor Cyan
Write-Host "Total items processed: $processedItems" -ForegroundColor White
Write-Host "Successfully extracted: $successItems" -ForegroundColor Green
Write-Host "Failed or skipped: $failedItems" -ForegroundColor Yellow
Write-Host "`nText extraction completed. Files saved to: $destinationRoot" -ForegroundColor Green
