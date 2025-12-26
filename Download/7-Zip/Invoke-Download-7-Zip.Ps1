<#
.SYNOPSIS
    Downloads 7-Zip.
.DESCRIPTION
    This script downloads the latest version of 7-Zip and saves it to a specified directory
.NOTES
    FileName:    Invoke-Download-7-Zip.Ps1
    Author:      Dag Langfloen
    Curent version: 1.0.3
    Version:        1.0.0 - Initial version
                    1.0.1 - Fixed Changed folder structure   
                    1.0.2 - Added -UseBasicParsing to Invoke-WebRequest
                    1.0.3 - Changed folderstructiure for Downloads 

#>

# Folders 
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$ApplicationName = "7-Zip"
$SystemDrive = $env:SystemDrive
# DownloadFolder
$ApplicationFolderPath = Join-Path $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose
$DownloadFolder = Join-Path $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $ApplicationName

# Create the folder if it doesn't exist
If (!(Test-Path $DownloadFolder))
{
    try {
        New-Item -Path $ApplicationFolderPath -Name $ApplicationName -ItemType Directory -ErrorAction Stop
    }
    catch {
        $Error
    }
}

# Find the latest 7-Zip version from the download page
$DownloadPage = Invoke-WebRequest -UseBasicParsing -Uri "https://www.7-zip.org/download.html"

# Find the latest 7-Zip version from the download page
$VersionPattern = 'Download 7-Zip ([\d\.]+)'
$LatestVersion = ($DownloadPage.Content | Select-String -Pattern $VersionPattern).Matches[0].Groups[1].Value
Write-Host "Latest 7-Zip version: $LatestVersion"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $LatestVersion
If (!(Test-Path $VersionFolder))
{
    try {
        New-Item -Path $DownloadFolder -Name $LatestVersion -ItemType Directory -ErrorAction Stop
    }
    catch {
        $Error
    }
}

# Change the download folder to the version folder
$DownloadFolder = $VersionFolder

# Removing the . in the version
$VersionNoDots = $LatestVersion -replace '\.', ''

# Get the latest 7-Zip x64 installer URL from the official site
$DownloadPage = Invoke-WebRequest -UseBasicParsing -Uri "https://www.7-zip.org/download.html"
$InstallerUrl = ($DownloadPage.Links | Where-Object {
    $_.href -match "a/7z$VersionNoDots-x64\.exe"
}).href

# Prepend domain if needed
if ($InstallerUrl -notmatch "^https?://") {
    $InstallerUrl = "https://www.7-zip.org/$InstallerUrl"
}

Write-Host "Latest 7-Zip x64 installer URL: $InstallerUrl"

# DownloadPath
$DownloadPath = Join-Path $DownloadFolder "7z$LatestVersion-x64.exe"

# Test if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

# Download the installer from the URL
Invoke-WebRequest -UseBasicParsing -Uri $InstallerUrl -OutFile $DownloadPath
Write-Host "7-Zip installer downloaded to: $DownloadPath"