<#
.SYNOPSIS
    Downloads Visual Studio Code.  
.DESCRIPTION
    This script downloads the latest version of Visual Studio Code and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadVSCode.Ps1
    Version:     1.0.0
    Author:      Dag Langfloen

#>

# Downloadfolder for Visual Studio Code
$DownloadFolder = "C:\Applications\Visual Studio Code"

# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
}

# Get the latest version of Visual Studio Code from GitHub API
$latest = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/vscode/releases/latest"
$latestVersion = $latest.tag_name -replace '^v', ''
Write-Host "Latest Visual Studio Code version: $latestVersion"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $latestVersion
If (!(Test-Path $VersionFolder)) {
    New-Item -Path $DownloadFolder -Name $latestVersion -ItemType Directory | Out-Null
}

# Get the installer name
$ApplicationName = "VSCodeSetup-x64-$latestVersion.exe"
$DownloadPath = Join-Path $VersionFolder $ApplicationName   

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

# Installer name
$InstallerName = "VSCodeSetup-x64-$latestVersion.exe"
$DownloadPath = Join-Path $VersionFolder $InstallerName

# Download URL for the installer
# Use the official static URL for the latest stable system installer (x64)
$installerUrl = "https://update.code.visualstudio.com/latest/win32-x64/stable"

# Download the installer
Invoke-WebRequest -Uri $installerUrl -OutFile $DownloadPath
Write-Host "Downloaded Visual Studio Code system installer to $DownloadPath"