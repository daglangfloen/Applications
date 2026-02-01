<#
.SYNOPSIS
    Downloads and installs GitHub Desktop from the official GitHub repository.
.DESCRIPTION
    This script downloads the latest version of GitHub Desktop from GitHub's releases and saves it to the Downloads folder.
.NOTES
    FileName:    Invoke-DownloadGitHubDesktop.Ps1
    Author:      Dag Langfloen
    Current version: 1.0
    Version:     1.0 - Initial version
                 

#>

# Variables
$ApplicationFolderName = "Applications"
$DownloadFolder = "Downloads"
$Developer = "GitHub"
$ApplicationName = "GitHub Desktop"
$SystemDrive = $env:SystemDrive

# Construct the download path
$DownloadPath = Join-Path $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $DownloadFolder | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName

# Create the download directory if it doesn't exist
If (!(Test-Path $DownloadPath)) {
    try {
        New-Item -Path $DownloadPath -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $DownloadPath"
        $Error
    }
}

# Define the URL for the latest GitHub Desktop release
$GitHubReleasesUrl = "https://api.github.com/repos/desktop/desktop/releases/latest"
try {
    # Fetch the latest release information from GitHub API
    $ReleaseInfo = Invoke-RestMethod -Uri $GitHubReleasesUrl -ErrorAction Stop

    # Extract the download URL for the Windows installer
    $Asset = $ReleaseInfo.assets | Where-Object { $_.name -like "*GitHubDesktopSetup-x64.exe" }
    If ($null -eq $Asset) {
        throw "No suitable installer found in the latest release."
    }
    $DownloadUrl = $Asset.browser_download_url
    $FileName = $Asset.name
    $DestinationFilePath = Join-Path $DownloadPath -ChildPath $FileName

    # Download the installer
    Write-Host "Downloading GitHub Desktop from $DownloadUrl to $DestinationFilePath"
    Invoke-WebRequest -UseBasicParsing -Uri $DownloadUrl -OutFile $DestinationFilePath -ErrorAction Stop

    Write-Host "Download completed successfully."
}
catch {
    Write-Host "An error occurred while downloading GitHub Desktop: $_"
}

# Getting the version of the downloaded installer
if (Test-Path $DestinationFilePath) {
    $FileVersionInfo = Get-Item $DestinationFilePath | Select-Object -ExpandProperty VersionInfo
    $Version = $FileVersionInfo.FileVersion
    Write-Host "Downloaded GitHub Desktop version: $Version"
} else {
    Write-Host "The installer file was not found at $DestinationFilePath"
}

# Create a version folder if it doesn't exist
$VersionFolderPath = Join-Path $DownloadPath -ChildPath $Version
If (!(Test-Path $VersionFolderPath)) {
    try {
        New-Item -Path $VersionFolderPath -ItemType Directory | Out-Null -ErrorAction Stop    
        # Move the downloaded installer to the version folder
        $FinalInstallerPath = Join-Path $VersionFolderPath -ChildPath $FileName
        try {
            Move-Item -Path $DestinationFilePath -Destination $FinalInstallerPath -ErrorAction Stop
            Write-Host "Moved installer to version folder: $FinalInstallerPath"
        }
        catch {
            Write-Host "An error occurred while moving the installer: $_"
        }
    }
    catch {
        Write-Host "Error creating version folder: $VersionFolderPath"
        $Error
    }
}
else {
    Write-Host "Version folder already exists: $VersionFolderPath"
    Write-Host "Processing of the downloaded installer is complete."
    # Removing the downloaded installer as it is already in the version folder
    if (Test-Path $DestinationFilePath) {
        Remove-Item -Path $DestinationFilePath -ErrorAction SilentlyContinue
    }
}

Write-Host "Script execution completed."