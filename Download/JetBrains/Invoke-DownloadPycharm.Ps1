<#
.SYNOPSIS
    Downloads PyCharm from JetBrains.
.DESCRIPTION
    This script downloads the latest PyCharm installer from JetBrains website and saves it to a specified Directory
.NOTES
    FileName:    Invoke-DownloadPycharm.Ps1
    Author:      Dag Langfloen
    Current version: 1.0
    Version:     1.0 - Initial version

#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$Developer = "JetBrains"
$ApplicationName = "PyCharm"
$SystemDrive = $env:SystemDrive

# Download the latest PyCharm installer (x64)
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

# Get latest release info from JetBrains website
$jetBrainsReleasesUrl = "https://data.services.jetbrains.com/products/releases?code=PCP&latest=true&type=release"
try {
    $releaseInfo = Invoke-RestMethod -Uri $jetBrainsReleasesUrl -ErrorAction Stop
    $latestVersion = $releaseInfo.PCP[0].version
    Write-Host "Latest PyCharm version: $latestVersion"

    # Construct download URL
    $downloadUrl = "https://download.jetbrains.com/python/pycharm-$latestVersion.exe"
    $fileName = "pycharm-$latestVersion.exe"
    $destinationFilePath = Join-Path $DownloadFolder -ChildPath $fileName

    # Download the installer
    Write-Host "Downloading PyCharm from $downloadUrl to $destinationFilePath"
    Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -OutFile $destinationFilePath -ErrorAction Stop

    Write-Host "Download completed successfully."
}
catch {
    Write-Host "An error occurred while downloading PyCharm: $_"
}

# Get the version of the downloaded installer
try {
    $FileVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($destinationFilePath)
    Write-Host "Downloaded PyCharm version: $($FileVersionInfo.ProductVersion)"
}
catch {
    Write-Host "An error occurred while retrieving the file version: $_"
}

# Create a folder for the version if needed and ifnot needed remove the downloaded installer
$VersionFolder = Join-Path $DownloadFolder -ChildPath $latestVersion

If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $VersionFolder -ItemType Directory | Out-Null -ErrorAction Stop    
        Move-Item -Path $destinationFilePath -Destination $VersionFolder -ErrorAction Stop
        Write-Host "Moved installer to version folder: $VersionFolder"
    }
    catch {
        Write-Host "An error occurred while creating version folder or moving the installer: $_"
    }
}
else {
    try {
        Remove-Item -Path $destinationFilePath -ErrorAction Stop
        Write-Host "Version folder already exists. Removed downloaded installer: $destinationFilePath"
    }
    catch {
        Write-Host "An error occurred while removing the downloaded installer: $_"
    }
}