<#
.SYNOPSIS
    Downloads the latest version of Adobe Acrobat DC Professional.
.DESCRIPTION
    This script downloads the Adobe Acrobat DC Professional installer to the specified directory.
.NOTES
    Script Name: Invoke-DownloadAdobeAcrobatPro.ps1
    Author: Dag Langfloen
    Current Version: 1.0.2
    Version History:
        1.0.0 - Initial release
        1.0.1 - Added -UseBasicParsing to Invoke-WebRequest
        1.0.2 - Changed folderstructiure for Downloads 
#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$ApplicationName = "Adobe Acrobat DC" 
$Sku = "Pro" 
$Developer = "Adobe Inc"
$SystemDrive = $env:SystemDrive

# DownloadFolder
$DownloadFolder = Join-Path $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName  | Join-Path -ChildPath $Sku

# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    try {
        New-Item -Path $DownloadFolder -ItemType Directory | Out-Null -ErrorAction Stop
    }
    catch {
        Write-Host "Error creating folder: $DownloadFolder" -ForegroundColor Red
        $Error
    }
}

# Import Evergreen module and update it
Import-Module Evergreen -ErrorAction Stop
Update-Module -Name Evergreen -Force
Update-Evergreen -Force

# Get the download URL for the latest version of Adobe Acrobat DC
# Find-EvergreenApp -Name 'AdobeAcrobat'
$InstallerUrl = Get-EvergreenApp -Name 'AdobeAcrobatProStdDC'  | Where-Object { $_.Architecture -eq 'x64' -and $_.Sku -eq $Sku } | Select-Object -ExpandProperty URI
Write-Host "Adobe Acrobat DC download URL: $InstallerUrl"

# Get the FileName from the URL
$FileName = ($InstallerUrl -split "/")[-1]
Write-Host "Installer file name: $FileName"
$InstallerPath = Join-Path -Path $DownloadFolder -ChildPath $FileName

# Check if the installer already exists, and if so, delete it
Write-Host "Checking for existing installer at: $InstallerPath"
If (Test-Path $InstallerPath) {
    try {
        Remove-Item $InstallerPath -Force -ErrorAction Stop
        Write-Host "Removed existing installer: $InstallerPath"
    }
    catch {
        Write-Host "Error removing existing installer: $InstallerPath"
        $Error
    }
} else {
    Write-Host "No existing installer found at: $InstallerPath"
    }

# Notify about downloading
Write-Host "Downloading Adobe Acrobat DC from: $InstallerUrl"
# Download the Adobe Acrobat DC installer
try {
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -UseBasicParsing -Uri $InstallerUrl -OutFile $InstallerPath -ErrorAction Stop
    Write-Host "Downloaded Adobe Acrobat DC installer to: $InstallerPath"
} 
catch {
    Write-Host "Error downloading Adobe Acrobat DC installer from: $InstallerUrl"
    $Error
}

# Unblock the zip file
Write-Host "Unblocking the installer zip file: $InstallerPath"
Unblock-File -Path $InstallerPath

# Extract the zip-file to the download folder
Write-Host "Extracting Adobe Acrobat DC installer to: $DownloadFolder"
try
{
    Expand-Archive -Path $InstallerPath -DestinationPath $DownloadFolder -Force -ErrorAction Stop
    Write-Host "Extracted Adobe Acrobat DC installer to: $DownloadFolder"
}
catch
{
    Write-Host "Error extracting Adobe Acrobat DC installer to: $DownloadFolder"
    $Error
}

# Remove the zip file after extraction
Write-Host "Removing installer zip file: $InstallerPath"
try {
    Remove-Item $InstallerPath -Force -ErrorAction Stop
    Write-Host "Removed installer zip file: $InstallerPath"
}
catch {
    Write-Host "Error removing installer zip file: $InstallerPath"
    $Error
}

# Find the version number from the .msp file in the extracted folder
Write-Host "Finding the  version number from the .msp file in: $DownloadFolder"
$MspFile = Get-ChildItem -Path $DownloadFolder -Filter "*.msp" -Recurse | Select-Object -First 1

# Get the version from the .msp filename by removing the 15 first characters and the .msp extension
$Version = $MspFile.BaseName.Substring(15)

# Add a . after the first two numbers, and after the next three numbers to format it as X.XX.XXX
$Version = $Version.Insert(2, ".").Insert(6, ".")
Write-Host "Downloaded Adobe Acrobat DC version: $Version"

# Create a version folder if it doesn't exist
Write-Host "Creating version folder: $VersionFolder"
$VersionFolder = Join-Path -Path $DownloadFolder -ChildPath $Version
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $VersionFolder -ItemType Directory | Out-Null -ErrorAction Stop
        Write-Host "Created version folder: $VersionFolder"
    }
    catch {
        Write-Host "Error creating version folder: $VersionFolder" -ForegroundColor Red
        $Error
    }
} else {
    Write-Host "Version folder already exists: $VersionFolder"
    Write-Host "Exiting script to avoid overwriting existing version folder." -ForegroundColor Yellow

}

# Move the extracted files to the version folder
Write-Host "Moving extracted files to version folder: $VersionFolder"
try {
    Move-Item -Path (Join-Path -Path $DownloadFolder -ChildPath "A*") -Destination $VersionFolder -Force -ErrorAction Stop
    Write-Host "Moved extracted files to version folder: $VersionFolder"
}
catch {
    Write-Host "Error moving extracted files to version folder: $VersionFolder" -ForegroundColor Red
    $Error
}
Write-Host "Adobe Acrobat DC download and extraction completed successfully." -ForegroundColor Green