# PowerShell script to move FirstPlugin.addin to Revit Addins folder
# File: MoveAddin.ps1

# Define source and destination paths
$sourcePath = "C:\Users\jensd\Documents\JD-Tegnogbyg\Plugin\Git\RevitJDByg\Addin_filer\FirstPlugin.addin"
$destinationPath = "C:\Users\jensd\AppData\Roaming\Autodesk\Revit\Addins\2024\FirstPlugin.addin"

# Create the destination directory if it doesn't exist
$destinationFolder = Split-Path -Path $destinationPath -Parent
if (-not (Test-Path -Path $destinationFolder)) {
    Write-Host "Creating destination directory: $destinationFolder"
    New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
}

# Check if source file exists
if (Test-Path -Path $sourcePath) {
    # Check if destination file already exists
    if (Test-Path -Path $destinationPath) {
        Write-Host "Destination file already exists. Replacing it..."
        Remove-Item -Path $destinationPath -Force
    }
    
    # Copy the file
    Copy-Item -Path $sourcePath -Destination $destinationPath -Force
    
    # Verify the file was copied successfully
    if (Test-Path -Path $destinationPath) {
        Write-Host "File successfully moved to $destinationPath"
    } else {
        Write-Host "Error: Failed to copy the file to the destination."
    }
} else {
    Write-Host "Error: Source file not found at $sourcePath"
}
