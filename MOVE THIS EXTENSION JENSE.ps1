# PowerShell script to copy pyRevit extension to test folder (P:\JDBYG REVIT)

# Get the current script directory
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define source and destination paths for testing
$sourcePath = Join-Path $scriptDirectory "pyrevit_export_materials.extension"
$testDestinationPath = "C:\Users\JENDAM\AppData\Roaming\pyRevit-Master\extensions"

Write-Host "=== pyRevit Extension Test Copy ===" -ForegroundColor Cyan
Write-Host "This is a TEST run - copying to: $testDestinationPath" -ForegroundColor Yellow
Write-Host ""

# Check if source folder exists
if (-Not (Test-Path $sourcePath)) {
    Write-Error "Source folder not found: $sourcePath"
    Write-Host "Please ensure the 'pyrevit_export_materials.extension' folder is in the same directory as this script."
    exit 1
}

# Check if test destination exists
if (-Not (Test-Path $testDestinationPath)) {
    Write-Host "Test destination folder not found: $testDestinationPath" -ForegroundColor Red
    Write-Host "Please ensure the P:\JDBYG REVIT folder exists on your machine."
    exit 1
}

# Define the full destination path for the extension
$fullDestinationPath = Join-Path $testDestinationPath "pyrevit_export_materials.extension"

# Check if extension already exists at test destination
if (Test-Path $fullDestinationPath) {
    Write-Host "Extension already exists at test destination." -ForegroundColor Yellow
    $response = Read-Host "Do you want to overwrite it? (Y/N)"
    if ($response -ne 'Y' -and $response -ne 'y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
    
    # Remove existing extension
    try {
        Remove-Item -Path $fullDestinationPath -Recurse -Force
        Write-Host "Existing test extension removed." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to remove existing test extension: $_"
        exit 1
    }
}

# Copy the extension folder
try {
    Write-Host "Copying extension for testing..." -ForegroundColor Blue
    Write-Host "From: $sourcePath" -ForegroundColor Gray
    Write-Host "To: $fullDestinationPath" -ForegroundColor Gray
    
    Copy-Item -Path $sourcePath -Destination $testDestinationPath -Recurse -Force
    
    Write-Host ""
    Write-Host "Test copy completed successfully!" -ForegroundColor Green
    Write-Host "Extension copied to: $fullDestinationPath" -ForegroundColor Gray
    
    # Display folder contents for verification
    Write-Host ""
    Write-Host "Verifying copy - Contents of test destination:" -ForegroundColor Cyan
    if (Test-Path $fullDestinationPath) {
        $items = Get-ChildItem -Path $fullDestinationPath -Recurse
        Write-Host "Total items copied: $($items.Count)" -ForegroundColor Green
        
        # Show first few items
        Write-Host "Sample contents:" -ForegroundColor Cyan
        $items | Select-Object -First 10 | ForEach-Object {
            Write-Host "  $($_.Name)" -ForegroundColor White
        }
        
        if ($items.Count -gt 10) {
            Write-Host "  ... and $($items.Count - 10) more items" -ForegroundColor Gray
        }
    }
}
catch {
    Write-Error "Test copy failed: $_"
    exit 1
}

Write-Host ""
Write-Host "TEST COMPLETE" -ForegroundColor Magenta
Write-Host "The extension has been copied to the test folder." -ForegroundColor White
Write-Host "Original extension remains in the script directory." -ForegroundColor White
Write-Host ""
Read-Host "Press Enter to exit"
