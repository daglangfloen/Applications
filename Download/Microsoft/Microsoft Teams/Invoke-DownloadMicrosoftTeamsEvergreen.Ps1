<#
.SYNOPSIS
    Downloads the latest Microsoft Teams Evergreen installer.
.DESCRIPTION
    This script downloads the latest version of Microsoft Teams Evergreen and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadMicrosoftTeamsEvergreen.Ps1
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
$Developer = "Microsoft Corporation"
$ApplicationName = "Microsoft Teams"
$SystemDrive = $env:SystemDrive

# Downloadfolder for Microsoft Teams Evergreen
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName

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

# Download the latest Teams installer
$TeamsUrl = "https://go.microsoft.com/fwlink/?linkid=2243204"
$TeamsInstaller = "MSTeamsSetup.exe"
$TeamsPath = Join-Path $DownloadFolder $TeamsInstaller

# Check if file already exists and delete if it does
If (Test-Path $TeamsPath) {
    try {
        Remove-Item $TeamsPath -Force -ErrorAction Stop    
    }
    catch {
        Write-Host "Error deleting file: $TeamsPath"
        $Error
    }
    
}

# Download the installer
try {
    Invoke-WebRequest -UseBasicParsing -Uri $TeamsUrl -OutFile $TeamsPath  -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Microsoft Teams from: $TeamsUrl" -ForegroundColor Red
    $Error
}


# Get version info from the downloaded installer
$version = (Get-Item $TeamsPath).VersionInfo.ProductVersion
Write-Output "Latest Teams Version: $version"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $version 
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $DownloadFolder -Name $version -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
    
}
# Move the installer to the version folder
$DestinationPath = Join-Path $VersionFolder $TeamsInstaller
try {
    Move-Item -Path $TeamsPath -Destination $DestinationPath -Force -ErrorAction Stop    
}
catch {
    Write-Host "Error moving file to: $DestinationPath" -ForegroundColor Red
    $Error
}
Write-Host "$ApplicationName downloaded to: $DestinationPath"