<#
.SYNOPSIS
    Downloads the latest version of the Powershell App Deployment Toolkit (PSADT).
.DESCRIPTION
    This script downloads the latest version of the Powershell App Deployment Toolkit (PSADT) from GitHub and saves it to a specified directory.
    The script creates a folder named "PSAppDeployToolkit" in the "C:\Applications" directory and extracts the contents of the downloaded zip file into this folder.
    
    The script is based on Patch My Pcs origial script: https://patchmypc.com/powershell-app-deployment-toolkit-download-script
.NOTES
    FileName:    Invoke-DownloadLatestPSADT.Ps1
    Author:      Dag Langfloen
    Curent version: 1.0.0
    Version:        1.0.0 - Initial version
                    1.0.1 - Changed folder structure to support Developer, and some error handling
                    1.0.2 - Added -UseBasicParsing to Invoke-WebRequest
                    1.0.3 - Changed folderstructiure for Downloads
    

#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$ApplicationName = "PSAppDeployToolkit"
$SystemDrive = $env:SystemDrive

# Downloadfolder for PSADT
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
    try {
        New-Item -Path $DownloadFolder -Name $latestVersion -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
    
}

# If the version folder already exists, delete the content
If (Test-Path $VersionFolder) {
    try {
        Remove-Item -Path (Join-Path $VersionFolder "*") -Recurse -Force -ErrorAction Stop    
    }
    catch {
        Write-Host "Error deleting content in folder: $VersionFolder"
        $Error
    }
    
}

# Download path
$DownloadPath = Join-Path $DownloadFolder $filenamePatternMatch

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

## Download to the file
try {
    Invoke-WebRequest -UseBasicParsing -Uri $psadtDownloadUri -Out $DownloadPath -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading PSADT from: $psadtDownloadUri" -ForegroundColor Red
    $Error
}

# Unblock the downloaded file
Unblock-File -Path $DownloadPath

## Extract the contents of the zip file to a subfolder under the current user's Downloads folder
try {
    Expand-Archive -Path $DownloadPath -DestinationPath $VersionFolder -Force -ErrorAction Stop    
}
catch {
    Write-Host "Error extracting PSADT zip file: $DownloadPath" -ForegroundColor Red
    $Error
}
Write-Host ("File: {0} extracted to Path: {1}" -f $psadtDownloadUri, $VersionFolder) -ForegroundColor Green

# Remove the zip file after extraction
try {
    Remove-Item $DownloadPath -Force -ErrorAction Stop    
}
catch {
    Write-Host "Error deleting zip file: $DownloadPath" -ForegroundColor Red
    $Error
}