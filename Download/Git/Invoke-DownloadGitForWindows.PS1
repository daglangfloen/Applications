<#
.SYNOPSIS
    Downloads the latest Git for Windows.
.DESCRIPTION
    This script downloads the latest version of Git for Windows and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadGitForWindows.PS1
    Author:      Dag Langfloen
    Version:     1.0.0
                 1.0.1 - Fixed Changed folder structure and added error handling
    
#>

# Folders
$ApplicationFolderName = "Applications"
$Developer = "Git"
$ApplicationName = "Git for Windows"
$SystemDrive = $env:SystemDrive
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName

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

# Find the latest Git for Windows release from GitHub API
$latest = Invoke-RestMethod -Uri "https://api.github.com/repos/git-for-windows/git/releases/latest"

# Latest version tag without .windows.1 suffix and v prefix 
$latestVersion = $latest.tag_name -replace '\.windows\.1$' -replace '^v', ''
Write-Host "Latest Git for Windows version: $latestVersion"

$VersionFolder = Join-Path $DownloadFolder $latestVersion
# Create download folder if it doesn't exist
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $VersionFolder -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
    
}

# Get the installer name
$ApplicationName = "Git-$latestVersion-64-bit.exe"
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

# Download the installer
try {
    Invoke-WebRequest -Uri "https://git-scm.com/download/win" -OutFile $DownloadPath -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Git for Windows from: https://git-scm.com/download/win" -ForegroundColor Red
    $Error
}
Write-Host "Downloaded Git for Windows installer to: $DownloadPath"