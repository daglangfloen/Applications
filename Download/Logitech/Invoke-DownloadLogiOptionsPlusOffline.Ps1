<#
.SYNOPSIS
    Downloads Logitech Options Plus offline installer.
.DESCRIPTION
    This script downloads the Logitech Options Plus offline installer and extracts it to the specified destination folder.
.NOTES
    File Name: Invoke-DownloadLogiOptionsPlusOffline.Ps1
    Author: Dag Langfloen
    Current Version: 1.1
    Version history: 1.0 : First release
                     1.1 : Added -UseBasicParsing to Invoke-WebRequest

#>

# Folders for Logitech Options Plus
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$Developer = "Logitech" 
$LogiOptionsPlusFolderName = "Logi Options Plus"
$SystemDrive = $env:SystemDrive
$ApplicationsFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose
$LogiOptionsPlusFolder = Join-Path $ApplicationsFolder -ChildPath $Developer | Join-Path -ChildPath $LogiOptionsPlusFolderName
$DownloadFolder = Join-Path $LogiOptionsPlusFolder -ChildPath "Offline"

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

# URL for Logitech Options Plus offline installer
$LogiOptionsPlusUrl = "https://download01.logi.com/web/ftp/pub/techsupport/optionsplus/logioptionsplus_installer_offline.exe"
$LogiOptionsPlusInstaller = Join-Path $DownloadFolder "logioptionsplus_installer_offline.exe"

# Download the installer if it doesn't exist
If (!(Test-Path $LogiOptionsPlusInstaller)) {
    try {
        Write-Output "Downloading Logitech Options Plus offline installer..."
        Invoke-WebRequest -UseBasicParsing -Uri $LogiOptionsPlusUrl -OutFile $LogiOptionsPlusInstaller -ErrorAction Stop
        Write-Output "Download completed: $LogiOptionsPlusInstaller"
    }
    catch {
        Write-Error "An error occurred while downloading the installer: $_"
    }
} else {
    Write-Output "Logitech Options Plus offline installer already exists: $LogiOptionsPlusInstaller"
}

# Get the version number from the metadata of the installer
$Version = (Get-Item $LogiOptionsPlusInstaller).VersionInfo.ProductVersion
Write-Output "Logitech Options Plus Version: $Version"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $Version
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $DownloadFolder -Name $Version -ItemType Directory | Out-Null -ErrorAction Stop
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
}
else {
    Write-Host "Version folder already exists: $VersionFolder"
    # Remove the downloaded installer to avoid confusion
    Remove-Item $LogiOptionsPlusInstaller -Force -ErrorAction SilentlyContinue
    Exit

}

# Move the installer to the version folder
Move-Item -Path $LogiOptionsPlusInstaller -Destination $VersionFolder -Force -ErrorAction Stop

# Remove the installer if it still exists in the download folder
If (Test-Path $LogiOptionsPlusInstaller) {
    try {
        Remove-Item $LogiOptionsPlusInstaller -Force -ErrorAction SilentlyContinue    
    }
    catch {
        Write-Host "Error deleting file: $LogiOptionsPlusInstaller" -ForegroundColor Red
        $Error
    }
}