<#
.SYNOPSIS
    Downloads Google Chrome Enterprise.
.DESCRIPTION
    This script downloads the latest version of Google Chrome Enterprise and saves it to a specified Directory
.NOTES
    FileName:    Invoke-DownloadChromeEnterprise.Ps1
    Author:      Dag Langfloen
    Version:     1.0.0
    
#>

# File and Folders
$Application = "GoogleChromeStandaloneEnterprise64.msi"
$ApplicationFolderName = "Applications"
$Developer = "Google"
$ApplicationName = "Google Chrome Enterprise"
$SystemDrive = $env:SystemDrive
# DownloadFolder
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName
$DownloadPath = Join-Path $DownloadFolder $Application

# URL to download the MSI installer
$DownlloadUrl = "https://dl.google.com/edgedl/chrome/install/$application"

# Create the folder if it doesn't exist
If (!(Test-Path $DownloadFolder))
{
    try {
        New-Item -Path $DownloadFolder -ItemType Directory | Out-Null -ErrorAction Stop
    }
    catch {
        $Error
    }
}

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    try {
        Remove-Item $DownloadPath -Force -ErrorAction Stop    
    }
    catch {
        Write-Host "Error deleting file: $DownloadPath" -ForegroundColor Red
        $Error
    }
    
}

# Download the MSI installer (if not already downloaded)
try {
    Invoke-WebRequest -Uri $DownlloadUrl -OutFile $DownloadPath -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Google Chrome Enterprise from: $DownlloadUrl" -ForegroundColor Red
    $Error
}

Write-Host "Downloaded to: $DownloadPath"

# Get the version from the MSI metadata
# by using Shell.Application COM object to read the file properties
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace((Get-Item $DownloadPath).DirectoryName)
$file = $folder.ParseName((Get-Item $DownloadPath).Name)
$subject = $folder.GetDetailsOf($file, 24)
$versionRegex = "\d+(.\d+){2,3}"
$version = ($subject | Select-String -Pattern $versionRegex).Matches.Value

Write-Host "Version: $version" -ForegroundColor Green

# Create subfolder for the version and move the file there
$VersionFolder = Join-Path $DownloadFolder $version
If (!(Test-Path $VersionFolder))
{
    try {
        New-Item -Path $DownloadFolder -Name $version -ItemType Directory -ErrorAction Stop
    }
    catch {
        $Error
    }
}
$NewDownloadPath = Join-Path $VersionFolder $Application

# If the file already exists in the version folder, remove it
If (Test-Path $NewDownloadPath) {
    try {
        Remove-Item $NewDownloadPath -Force -ErrorAction Stop    
    }
    catch {
        Write-Host "Error deleting file: $NewDownloadPath" -ForegroundColor Red
        $Error
    }
    
}

# Now move the file
try {
    Move-Item -Path $DownloadPath -Destination $NewDownloadPath -Force -ErrorAction Stop    
}
catch {
    Write-Host "Error moving file to: $NewDownloadPath" -ForegroundColor Red
    $Error
}
Write-Host "Application: $ApplicationName Moved to: $NewDownloadPath" -ForegroundColor Green