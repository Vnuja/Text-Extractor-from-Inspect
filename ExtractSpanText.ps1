# Define the path to the CSV file
$csvPath = "Descriptions.csv"

# Define the output folder
$outputFolder = "Descriptions"

# Create the output folder if it doesn't exist
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
}

# Import the CSV file
$csvData = Import-Csv -Path $csvPath

# Loop through each row in the CSV
foreach ($row in $csvData) {
    $fileName = $row.'Text File Name'
    $description = $row.'Main Description'

    # Extract text between <span> tags using regex
    $spanText = [regex]::Matches($description, '<span[^>]*>(.*?)</span>') | ForEach-Object { $_.Groups[1].Value }

    # Join the extracted text into a single string
    $extractedText = $spanText -join "`n"

    # Define the output file path
    $outputFilePath = Join-Path -Path $outputFolder -ChildPath "$fileName.txt"

    # Write the extracted text to the output file
    Set-Content -Path $outputFilePath -Value $extractedText
}

Write-Host "Extraction complete. Files saved in the '$outputFolder' folder."