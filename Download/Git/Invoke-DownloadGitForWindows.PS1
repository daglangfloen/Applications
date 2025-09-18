<#
.SYNOPSIS
    Downloads the latest Git for Windows.
.DESCRIPTION
    This script downloads the latest version of Git for Windows and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadGitForWindows.PS1
    Version:     1.0.0
    Author:      Dag Langfloen

#>

# Downloadfolder for Git for Windows
$DownloadFolder = "C:\Applications\Git"

# Find the latest Git for Windows release from GitHub API
$latest = Invoke-RestMethod -Uri "https://api.github.com/repos/git-for-windows/git/releases/latest"

# Latest version tag without .windows.1 suffix and v prefix 
$latestVersion = $latest.tag_name -replace '\.windows\.1$' -replace '^v', ''
Write-Host "Latest Git for Windows version: $latestVersion"

$VersionFolder = Join-Path $DownloadFolder $latestVersion
# Create download folder if it doesn't exist
If (!(Test-Path $VersionFolder)) {
    New-Item -Path $VersionFolder -ItemType Directory | Out-Null
}

# Get the installer name
$ApplicationName = "Git-$latestVersion-64-bit.exe"
$DownloadPath = Join-Path $VersionFolder $ApplicationName

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}   

# Download URL for the installer
Invoke-WebRequest -Uri "https://git-scm.com/download/win" -OutFile $DownloadPath 