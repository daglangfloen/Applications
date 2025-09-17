<#
.SYNOPSIS
    Downloads Notepad++.
.DESCRIPTION
    This script downloads the latest version of Notepad++ and saves it to a specified Directory
.NOTES
    FileName:    Invoke-DownloadNotepad++.Ps1
    Version:     1.0.0
    Author:      Dag Langfloen
#>


# Download the latest Notepad++ installer (x64)
$DownloadFolder = "C:\Applications\Notepad++"

# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
}

# Get latest release info from GitHub API
$apiUrl = "https://api.github.com/repos/notepad-plus-plus/notepad-plus-plus/releases/latest"
$release = Invoke-RestMethod -Uri $apiUrl

# Find the x64 installer asset
$asset = $release.assets | Where-Object { $_.name -like "*x64*.exe" } | Select-Object -First 1
$downloadUrl = $asset.browser_download_url

# Output the latest version tag
$latestVersion = $release.tag_name
Write-Host "Latest Notepad++ version: $latestVersion"

# Remove the 'v' prefix if present
if ($latestVersion -like "v*") {
    $latestVersion = $latestVersion.Substring(1)
}

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $latestVersion   
If (!(Test-Path $VersionFolder)) {
    New-Item -Path $DownloadFolder -Name $latestVersion -ItemType Directory | Out-Null
}

# Get the installer name
$InstallerName = "npp.$latestVersion.Installer.x64.exe"

# Download path
$DownloadPath = Join-Path $VersionFolder  $InstallerName

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

# Download the installer
Invoke-WebRequest -Uri $downloadUrl -OutFile $DownloadPath

Write-Host "Downloaded Notepad++ installer to: $DownloadPath"