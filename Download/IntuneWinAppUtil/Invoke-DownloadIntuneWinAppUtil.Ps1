<#
.SYNOPSIS
    Downloads the latest version of the Powershell App Deployment Toolkit (PSADT).
.DESCRIPTION
    This script downloads the latest version of the IntuneWinAppUtil from the official Microsoft GitHub repository and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadIntuneWinAppUtil.Ps1
    Author:      Dag Langfloen
    Curent version: 1.0.0
    Version:        1.0.0 - Initial version
                    1.0.1 - Added -UseBasicParsing to Invoke-WebRequest
                    1.0.2 - Changed folderstructiure for Downloads

#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$ApplicationName = "IntuneWinAppUtil"
$SystemDrive = $env:SystemDrive
$File = "IntuneWinAppUtil.exe"

# Downloadfolder for IntuneWinAppUtil
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $ApplicationName

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

# Download IntuneWinAppUtil from GitHub
$intuneWinAppUtilUri = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/blob/master/IntuneWinAppUtil.exe?raw=true"
$DownloadPath = Join-Path $DownloadFolder $File

try {
    Invoke-WebRequest -UseBasicParsing -Uri $intuneWinAppUtilUri -OutFile $DownloadPath -ErrorAction Stop
    Write-Host "IntuneWinAppUtil downloaded successfully to $DownloadPath" -ForegroundColor Green
}
catch {
    Write-Host "Error downloading IntuneWinAppUtil from $intuneWinAppUtilUri" -ForegroundColor Red
    $Error
}

# Find the latest version of IntuneWinAppUtil from GitHub
# Note: IntuneWinAppUtil does not have a formal versioning system exposed via API, so we will download the latest available from GitHub.
$versionInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/latest"
$LatestVersion = $versionInfo.tag_name
Write-Output "Latest IntuneWinAppUtil Version: $LatestVersion"

# Create a folder with the version number if it doesn't exist
$VersionFolder = Join-Path $DownloadFolder $LatestVersion
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $VersionFolder -ItemType Directory | Out-Null -ErrorAction Stop
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
}

# Move the downloaded file to the version folder if not already there
$VersionedDownloadPath = Join-Path $VersionFolder $File
try {
    Move-Item -Path $DownloadPath -Destination $VersionedDownloadPath -ErrorAction Stop
    Write-Host "IntuneWinAppUtil moved to version folder: $VersionedDownloadPath" -ForegroundColor Green    
}
catch {
    Write-Host "Error moving IntuneWinAppUtil to version folder: $VersionedDownloadPath" -ForegroundColor Red
    $Error
    Remove-Item -Path $DownloadPath -Force -ErrorAction Stop
}

Write-Host "Script completed." -ForegroundColor Green