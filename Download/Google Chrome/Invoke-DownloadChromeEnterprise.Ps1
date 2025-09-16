<#
.SYNOPSIS
    Downloads Google Chrome Enterprise.
.DESCRIPTION
    This script downloads the latest version of Google Chrome Enterprise and saves it to a specified Directory
.NOTES
    FileName:    Invoke-DownloadChromeEnterprise.Ps1
    Version:     1.0.0
    Author:      Dag Langfloen
#>

# DownloadFolder
$DownloadFolder = "C:\Applications\Google Chrome Enterprise"
$Application = "GoogleChromeStandaloneEnterprise64.msi"
$DownloadPath = Join-Path $DownloadFolder $Application

# URL to download the MSI installer
$DownlloadUrl = "https://dl.google.com/edgedl/chrome/install/$application"

# Create the folder if it doesn't exist
If (!(Test-Path $DownloadFolder))
{
    try {
        New-Item -Path "C:\Applications" -Name "Google Chrome Enterprise" -ItemType Directory -ErrorAction Stop
    }
    catch {
        $Error
    }
}

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

# Download the MSI installer (if not already downloaded)
Invoke-WebRequest -Uri $DownlloadUrl -OutFile $DownloadPath
Write-Host "Downloaded to: $DownloadPath"

# Get the version from the MSI metadata
# by using Shell.Application COM object to read the file properties
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace((Get-Item $DownloadPath).DirectoryName)
$file = $folder.ParseName((Get-Item $DownloadPath).Name)
$subject = $folder.GetDetailsOf($file, 24)
$versionRegex = "\d+(.\d+){2,3}"
$version = ($subject | Select-String -Pattern $versionRegex).Matches.Value

Write-Host "Version: $version"

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
    Remove-Item $NewDownloadPath -Force
}
# Now move the file
Move-Item -Path $DownloadPath -Destination $NewDownloadPath -Force