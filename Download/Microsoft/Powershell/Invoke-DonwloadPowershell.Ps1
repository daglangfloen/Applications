<#
.SYNOPSIS
    Downloads the latest version of Powershell.
.DESCRIPTION
    This script downloads the latest version of Powershell from the official Microsoft repository and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadPowershell.Ps1
    Author:      Dag Langfloen
    Curent version: 1.0.0
    Version:        1.0.0 - Initial version
#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$Vendor = "Microsoft Corporation"
$ApplicationName = "Powershell"
$SystemDrive = $env:SystemDrive

# Downloadfolder for Powershell
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $Vendor | Join-Path -ChildPath $ApplicationName

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

# Get the latest version of Powershell from GitHub API
$latest = Invoke-RestMethod -Uri "https://api.github.com/repos/PowerShell/PowerShell/releases/latest"
$latestVersion = $latest.tag_name -replace '^v', ''

Write-Host "Latest Powershell version: $latestVersion"

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
$ApplicationName = "PowerShell-$latestVersion-win-x64.msi"
$DownloadPath = Join-Path $VersionFolder $ApplicationName

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    try {
        Remove-Item -Path $DownloadPath -ErrorAction Stop
    }
    catch {
        Write-Host "Error deleting existing file: $DownloadPath"
        $Error
    }
}

# Download the installer
try {
    $asset = $latest.assets | Where-Object { $_.name -eq $ApplicationName } | Select-Object -First 1
    $downloadUrl = $asset.browser_download_url
    Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -OutFile $DownloadPath -ErrorAction Stop
    Write-Host "Downloaded Powershell version $latestVersion to $DownloadPath"
}
catch {
    Write-Host "Error downloading Powershell installer."
    $Error
}
Write-Host "Downloaded Powershell installer to: $DownloadPath"