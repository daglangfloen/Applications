<#
.SYNOPSIS
    Downloads Microsoft Edge.WebView2.
.DESCRIPTION
    This script downloads the latest version of Microsoft Edge WebView2 and saves it to a specified directory.  
.NOTES
    FileName:    Invoke-DownloadMicrosoftEdgeWebView2.Ps1
    Author:      Dag Langfloen
    Curent version: 1.0.0
    Version:        1.0.0 - Initial version
                    1.0.1 - Changed folder structure to support Developer, and some error handling

#>

# File and Folders
$Application = "MicrosoftEdgeWebView2RuntimeInstallerX64.msi"
$ApplicationFolderName = "Applications"
$Developer = "Microsoft Corporation"
$ApplicationName = "Microsoft Edge WebView2"
$SystemDrive = $env:SystemDrive

# Downloadfolder for Microsoft Edge WebView2
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

# Get the latest version of Microsoft Edge WebView2
$webviewInfo = Invoke-RestMethod -Uri "https://edgeupdates.microsoft.com/api/products?view=enterprise"
$webviewProduct = $webviewInfo | Where-Object { $_.Product -eq "Stable" }

# Get the latest version for windows x64
$latestStable = $webviewProduct.Releases | Where-Object { $_.Platform -eq "Windows" -and $_.Architecture -eq "x64" } | Select-Object -First 1
$LatestVersion = $latestStable.ProductVersion
Write-Output "Latest Stable WebView2 Version: $LatestVersion"

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

# DownloadPath
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

# Check if installer already exists and delete if it does
If (Test-Path $DownloadPath) {
    try {
        Write-Output "Installer already exists at $DownloadPath"
        Remove-Item $DownloadPath -Force -ErrorAction Stop    
    }
    catch {
        Write-Host "Error deleting file: $DownloadPath"
        $Error
    }
    
}

# Download URL for the installer
$webviewUrl = $latestStable.Artifacts | Select-Object -ExpandProperty Location
# Download the installer
try {
    Invoke-WebRequest -Uri $webviewUrl -OutFile $DownloadPath  -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Microsoft Edge WebView2 from: $webviewUrl" -ForegroundColor Red
    $Error
}
Write-Output "$ApplicationName downloaded to: $DownloadPath"    