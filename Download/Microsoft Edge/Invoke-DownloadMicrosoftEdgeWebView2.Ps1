<#
.SYNOPSIS
    Downloads Microsoft Edge.WebView2.
.DESCRIPTION
    This script downloads the latest version of Microsoft Edge WebView2 and saves it to a specified directory.  
.NOTES
    FileName:    Invoke-DownloadMicrosoftEdgeWebView2.Ps1
    Version:     1.0.0
    Author:      Dag Langfloen

#>

# Downloadfolder for Microsoft Edge WebView2
$DownloadFolder = "C:\Applications\Microsoft Edge WebView2"
# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
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
    New-Item -Path $VersionFolder -ItemType Directory | Out-Null
}

# DownloadPath
$InstallerName = "MicrosoftEdgeWebView2RuntimeInstallerX64.msi" 
$DownloadPath = Join-Path $VersionFolder $InstallerName
# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

# Check if installer already exists and delete if it does
If (Test-Path $DownloadPath) {
    Write-Output "Installer already exists at $DownloadPath"
    Remove-Item $DownloadPath -Force
}

# Download URL for the installer
$webviewUrl = $latestStable.Artifacts | Select-Object -ExpandProperty Location
Invoke-WebRequest -Uri $webviewUrl -OutFile $DownloadPath