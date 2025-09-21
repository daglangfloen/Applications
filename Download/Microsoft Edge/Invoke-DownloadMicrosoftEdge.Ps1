<#
.SYNOPSIS
    Downloads Microsoft Edge Enterprise.
.DESCRIPTION
    This script downloads the latest version of Microsoft Edge Enterprise and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadMicrosoftEdge.Ps1
    Author:      Dag Langfloen
    Current version: 1.0.2
    Version:         1.0.0  - Initial version
                     1.0.1     - Changed folder structure
                     1.0.2     - Added error handling
    
#>

# File and Folders
$Application = "MicrosoftEdgeEnterpriseX64.msi"
$ApplicationFolderName = "Applications"
$Developer = "Microsoft Corporation"
$ApplicationName = "Microsoft Edge"
$Channel = "Stable"
$SystemDrive = $env:SystemDrive

# Downloadfolder for Microsoft Edge
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName

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

# Get the latest version of Microsoft Edge Enterprise
$versionInfo = Invoke-RestMethod -Uri "https://edgeupdates.microsoft.com/api/products?view=enterprise"
$latestStable = $versionInfo | Where-Object { $_.Product -eq "Stable" } | Select-Object -ExpandProperty Releases # | Where-Object { $_.Channel == "Stable" } | Select-Object -First 1

# get the windows x64 version from the latest stable
$latestWindowsStable = $latestStable | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } | Select-Object -First 1
$LatestVersion = $latestWindowsStable.ProductVersion
Write-Output "Latest Stable Edge Version: $LatestVersion"

# Version folder
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

$DownloadPath = Join-Path $VersionFolder $Application

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

# Download URL for the installer
$edgeUrl = $latestWindowsStable.Artifacts | Select-Object -ExpandProperty Location # "https://msedgedownloads.azureedge.net/enterprise/LATEST_STABLE/msedgeenterprise_x64.msi"
try {
    Invoke-WebRequest -Uri $edgeUrl -OutFile $DownloadPath  -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Microsoft Edge Enterprise from: $edgeUrl" -ForegroundColor Red
    $Error
}
Write-Output "$ApplicationName downloaded to: $DownloadPath"