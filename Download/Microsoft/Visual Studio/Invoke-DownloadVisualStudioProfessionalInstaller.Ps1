<#
.SYNOPSIS
    Downloads the latest version of the Visual Studio Professional installer.
.DESCRIPTION
    This script downloads the latest version of Visual Studio Professional and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadVisualStudioProfessionalInstaller.Ps1
    Author:      Dag Langfloen
    Curent version: 1.0.0
    Version:        1.0.0 - Initial version
#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$Developer = "Microsoft Corporation"
$ApplicationName = "Visual Studio"
$edition = "Professional"
$SystemDrive = $env:SystemDrive

# Downloadfolder for Visual Studio Professional
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName | Join-Path -ChildPath $edition

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

# Download the latest Visual Studio Professional installer
$InstallerUrl = "https://aka.ms/vs/17/release/vs_professional.exe"
$InstallerFileName = "vs_professional.exe" 
$InstallerPath = Join-Path $DownloadFolder $InstallerFileName

# Check if file already exists and delete if it does
If (Test-Path $InstallerPath) {
    try {
        Remove-Item $InstallerPath -Force -ErrorAction Stop    
    }
    catch {
        Write-Host "Error deleting file: $InstallerPath"
        $Error
    }
    
}

# Download the installer
try {
    Invoke-WebRequest -UseBasicParsing -Uri $InstallerUrl -OutFile $InstallerPath  -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Visual Studio Community installer."
    $Error
}   
Write-Host "Downloaded to: $InstallerPath"

# Get the version from the installer metadata
$InstallerVersion = (Get-Item $InstallerPath).VersionInfo.FileVersion
Write-Host "Installer version: $InstallerVersion"

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $InstallerVersion
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $DownloadFolder -Name $InstallerVersion -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
    
}

# Move the installer to the version folder
$DestinationPath = Join-Path $VersionFolder $InstallerFileName
try {
    Move-Item -Path $InstallerPath -Destination $DestinationPath -Force -ErrorAction Stop    
}
catch {
    Write-Host "Error moving installer to version folder: $DestinationPath"
    $Error
}   
Write-Host "Moved installer to version folder: $DestinationPath"
# End of script