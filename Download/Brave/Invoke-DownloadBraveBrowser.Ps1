<#
.SYNOPSIS
    Downloads the Brave Browser installer.
.DESCRIPTION
    This script downloads the Brave Browser installer to the specified directory.
.NOTES
    Script Name: Invoke-DownloadBraveBrowser.ps1
    Author: Dag Langfloen
    Current Version: 1.0.0
    Version History:
        1.0.0 - Initial release
        1.0.1 - Added -UseBasicParsing to Invoke-WebRequest
        1.0.2 - Changed folderstructiure for Downloads
#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$ApplicationName = "Brave Browser"  
$Developer = "Brave"
$SystemDrive = $env:SystemDrive

# DownloadFolder
$DownloadFolder = Join-Path $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName

# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    try {
        New-Item -Path $DownloadFolder -ItemType Directory | Out-Null -ErrorAction Stop
    }
    catch {
        Write-Host "Error creating folder: $DownloadFolder"
        $Error
    }
}

# 
$InstallerUrl = "https://laptop-updates.brave.com/latest/win64"

# Check if the Brave Browser installer already exists
$InstallerFileName = "BraveBrowserSetup.exe"
 
$InstallerPath = Join-Path -Path $DownloadFolder -ChildPath $InstallerFileName

# Check if the installer already exists, and if so, delete it
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


# Download the Brave Browser installer
try {
    Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath -UseBasicParsing -ErrorAction Stop
    Write-Host "Downloaded Brave Browser installer to: $InstallerPath"    
}
catch {
    Write-Host "Error downloading Brave Browser installer from $InstallerUrl"
    $Error
}

# Get the version of the downloaded installer
$Version = (Get-Item $InstallerPath).VersionInfo.FileVersion
Write-Host "Downloaded Brave Browser installer version: $Version"

# Create a version folder
$VersionFolder = Join-Path -Path $DownloadFolder -ChildPath $Version
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $VersionFolder -ItemType Directory | Out-Null -ErrorAction Stop
        Write-Host "Created version folder: $VersionFolder"
    }
    catch {
        Write-Host "Error creating version folder: $VersionFolder"
        $Error
    }
} else {
    Write-Host "Version folder already exists: $VersionFolder"
}

# Move the installer to the version folder
$VersionedInstallerPath = Join-Path -Path $VersionFolder -ChildPath $InstallerFileName
try {
    Move-Item -Path $InstallerPath -Destination $VersionedInstallerPath -Force -ErrorAction Stop
    Write-Host "Moved installer to version folder: $VersionedInstallerPath"
}
catch {
    Write-Host "Error moving installer to version folder: $VersionedInstallerPath"
    $Error
}
Write-Host "Brave Browser installer download and organization complete."