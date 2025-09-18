<#
.SYNOPSIS
    Downloads the latest Microsoft Teams Evergreen installer.
.DESCRIPTION
    This script downloads the latest version of Microsoft Teams Evergreen and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadMicrosoftTeamsEvergreen.Ps1
    Version:     1.0.0
    Author:      Dag Langfloen

#>

# Downloadfolder for Microsoft Teams Evergreen
$DownloadFolder = "C:\Applications\Microsoft Teams"

# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
}

# Download the latest Teams installer
$TeamsUrl = "https://go.microsoft.com/fwlink/?linkid=2243204"
$TeamsInstaller = "MSTeamsSetup.exe"
$TeamsPath = Join-Path $DownloadFolder $TeamsInstaller

Invoke-WebRequest -Uri $TeamsUrl -OutFile $TeamsPath

# Get version info from the downloaded installer
$version = (Get-Item $TeamsPath).VersionInfo.ProductVersion
Write-Output "Latest Teams Version: $version"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $version 
If (!(Test-Path $VersionFolder)) {
    New-Item -Path $DownloadFolder -Name $version -ItemType Directory | Out-Null
}
# Move the installer to the version folder
$DestinationPath = Join-Path $VersionFolder $TeamsInstaller
Move-Item -Path $TeamsPath -Destination $DestinationPath -Force