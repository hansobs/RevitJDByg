# PowerShell script to move pyRevit extension to the extensions folder

# Get the current script directory
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define source and destination paths
$sourcePath = Join-Path $scriptDirectory "pyrevit_export_materials.extension"
$destinationPath = "C:\Users\JENDAM\AppData\Roaming\pyRevit-Master\extensions"

# Check if source folder exists
if (-Not (Test-Path $sourcePath)) {
    Write-Error "Source folder not found: $sourcePath"
    Write-Host "Please ensure the 'pyrevit_export_materials.extension' folder is in the same directory as this script."
    exit 1
}

# Create destination directory if it doesn't exist
if (-Not (Test-Path $destinationPath)) {
    Write-Host "Creating destination directory: $destinationPath"
    try {
        New-Item -ItemType Directory -Path $destinationPath -Force
        Write-Host "Destination directory created successfully."
    }
    catch {
        Write-Error "Failed to create destination directory: $_"
        exit 1
    }
}

# Define the full destination path for the extension
$fullDestinationPath = Join-Path $destinationPath "pyrevit_export_materials.extension"

# Check if extension already exists at destination
if (Test-Path $fullDestinationPath) {
    $response = Read-Host "Extension already exists at destination. Do you want to overwrite it? (Y/N)"
    if ($response -ne 'Y' -and $response -ne 'y') {
        Write-Host "Operation cancelled by user."
        exit 0
    }
    
    # Remove existing extension
    try {
        Remove-Item -Path $fullDestinationPath -Recurse -Force
        Write-Host "Existing extension removed."
    }
    catch {
        Write-Error "Failed to remove existing extension: $_"
        exit 1
    }
}

# Move the extension folder
try {
    Write-Host "Moving extension from: $sourcePath"
    Write-Host "To: $fullDestinationPath"
    
    Move-Item -Path $sourcePath -Destination $destinationPath -Force
    
    Write-Host "✓ Extension moved successfully!"
    Write-Host "The pyrevit_export_materials.extension is now installed at: $fullDestinationPath"
    Write-Host ""
    Write-Host "Next steps:"
    Write-Host "1. Restart Revit if it's currently running"
    Write-Host "2. The extension should appear in your pyRevit toolbar"
}
catch {
    Write-Error "Failed to move extension: $_"
    exit 1
}

# Optional: Pause to see results
Write-Host ""
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
