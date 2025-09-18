<#
.SYNOPSIS
    Downloads the latest version of the Powershell App Deployment Toolkit (PSADT).
.DESCRIPTION
    This script downloads the latest version of the Powershell App Deployment Toolkit (PSADT) from GitHub and saves it to a specified directory.
    The script creates a folder named "PSAppDeployToolkit" in the "C:\Applications" directory and extracts the contents of the downloaded zip file into this folder.
    
    The script is based on Patch My Pcs origial script: https://patchmypc.com/powershell-app-deployment-toolkit-download-script
.NOTES
    FileName:    Invoke-DownloadLatestPSADT.Ps1
    Version:     1.0.0
    Author:      Dag Langfloen

#>

# Downloadfolder for PSADT
$DownloadFolder = "C:\Applications\PSAppDeployToolkit"
# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
}

# Get the latest release info from GitHub API
## Download the latest release of PSAppDeployToolkit from GitHub
$githubRepository = "psappdeploytoolkit/psappdeploytoolkit"
$filenamePatternMatch = "PSAppDeployToolkit_Template_v4.zip"
$psadtReleaseUri = "https://api.github.com/repos/$githubRepository/releases/latest"
$psadtDownloadUri = ((Invoke-RestMethod -Method GET -Uri $psadtReleaseUri).assets | Where-Object name -like $filenamePatternMatch ).browser_download_url

# Find the version number
$psadtRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/PSAppDeployToolkit/PSAppDeployToolkit/releases/latest"
$latestVersion = $psadtRelease.tag_name
Write-Output "Latest PSADT Version: $latestVersion"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $latestVersion
If (!(Test-Path $VersionFolder)) {
    New-Item -Path $DownloadFolder -Name $latestVersion -ItemType Directory | Out-Null
}

# If the version folder already exists, delete the content
If (Test-Path $VersionFolder) {
    Remove-Item -Path (Join-Path $VersionFolder "*") -Recurse -Force
}

# Download path
$DownloadPath = Join-Path $DownloadFolder $filenamePatternMatch

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

## Download to the file
Invoke-WebRequest -Uri $psadtDownloadUri -Out $DownloadPath
## Remove any Zone.Identifier alternate data streams to unblock the file (if required)

# Unblock the downloaded file
Unblock-File -Path $DownloadPath

## Extract the contents of the zip file to a subfolder under the current user's Downloads folder
Expand-Archive -Path $DownloadPath -DestinationPath $VersionFolder -Force

Write-Host ("File: {0} extracted to Path: {1}" -f $psadtDownloadUri, $VersionFolder) -ForegroundColor Green