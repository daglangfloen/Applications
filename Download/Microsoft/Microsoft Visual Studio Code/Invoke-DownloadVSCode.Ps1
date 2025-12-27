<#
.SYNOPSIS
    Downloads Visual Studio Code.  
.DESCRIPTION
    This script downloads the latest version of Visual Studio Code and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadVSCode.Ps1
    Author:      Dag Langfloen
    Curent version: 1.0.0
    Version:        1.0.0 - Initial version
                    1.0.1 - Changed folder structure to support Developer, and some error handling
                    1.0.2 - Added -UseBasicParsing to Invoke-WebRequest
                    1.0.3 - Changed folderstructiure for Downloads
    

#>
# Folders
$ApplicationFolderName = "Applications" 
$Purpose = "Downloads"
$Developer = "Microsoft Corporation"
$ApplicationName = "Visual Studio Code"
$SystemDrive = $env:SystemDrive

# Downloadfolder for Visual Studio Code
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName

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

# Get the latest version of Visual Studio Code from GitHub API
$latest = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/vscode/releases/latest"
$latestVersion = $latest.tag_name -replace '^v', ''
Write-Host "Latest Visual Studio Code version: $latestVersion"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $latestVersion
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $DownloadFolder -Name $latestVersion -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
    
}

# Get the installer name
$ApplicationName = "VSCodeSetup-x64-$latestVersion.exe"
$DownloadPath = Join-Path $VersionFolder $ApplicationName   

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    try {
        Remove-Item $DownloadPath -Force -ErrorAction Stop    
    }
    catch {
        Write-Host "Error deleting file: $DownloadPath"
        $Error
    }
    
}

# Installer name
$InstallerName = "VSCodeSetup-x64-$latestVersion.exe"
$DownloadPath = Join-Path $VersionFolder $InstallerName

# Download URL for the installer
# Use the official static URL for the latest stable system installer (x64)
$installerUrl = "https://update.code.visualstudio.com/latest/win32-x64/stable"

# Download the installer
try {
    Invoke-WebRequest -UseBasicParsing -Uri $installerUrl -OutFile $DownloadPath -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Visual Studio Code from: $installerUrl" -ForegroundColor Red
    $Error
}
Write-Host "Downloaded Visual Studio Code system installer to $DownloadPath"