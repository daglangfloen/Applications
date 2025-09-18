<#
.SYNOPSIS
    Downloads Microsoft Edge Enterprise.
.DESCRIPTION
    This script downloads the latest version of Microsoft Edge Enterprise and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadMicrosoftEdge.Ps1
    Version:     1.0.0
    Author:      Dag Langfloen
#>

# Downloadfolder for Microsoft Edge
$DownloadFolder = "C:\Applications\Microsoft Edge"

# Create download folder if it doesn't exist
If (!(Test-Path $DownloadFolder)) {
    New-Item -Path $DownloadFolder -ItemType Directory | Out-Null
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
    New-Item -Path $VersionFolder -ItemType Directory | Out-Null
}

# Download the installer
$InstallerName = "MicrosoftEdgeEnterpriseX64.msi"   
$DownloadPath = Join-Path $VersionFolder $InstallerName

# Check if file already exists and delete if it does
If (Test-Path $DownloadPath) {
    Remove-Item $DownloadPath -Force
}

# Download URL for the installer
$edgeUrl = $latestWindowsStable.Artifacts | Select-Object -ExpandProperty Location # "https://msedgedownloads.azureedge.net/enterprise/LATEST_STABLE/msedgeenterprise_x64.msi"
Invoke-WebRequest -Uri $edgeUrl -OutFile $DownloadPath  