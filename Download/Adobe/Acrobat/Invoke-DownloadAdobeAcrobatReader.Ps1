<#
   .SYNOPSIS
   This script will download Adobe Acrobat Reader

.DESCRIPTION
   This script downloads Adobe Acrobat Reader and its customization wizard.
   The customization wizard must be installed as administrator. to be used.

.NOTES

    File Name           : Invoke-DownloadAdobeAcrobatReader.Ps1
    Author              : Dag Langfloen
    Current Version     : 1.3
    Version history     : 1.0  : First release
                          1.1  : Added -UseBasicParsing to Invoke-WebRequest
                          1.2 Changed folder structure
                          1.3 Changed the handling of the downloaded file
#>

# Folder for the downloadPath
$ApplicationFolderName = "Applications"
$Purpose = "Downloads"
$ApplicationName = "Adobe Acrobat DC" 
$Sku = "Reader" 
$Developer = "Adobe Inc"
$SystemDrive = $env:SystemDrive

# Download Path
$DownloadPath = Join-Path $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $Developer | Join-Path -ChildPath $ApplicationName | Join-Path -ChildPath $Sku

# Create Download folder if it does not exist
If (!(Test-Path -Path $DownloadPath)) {
    New-Item -Path $DownloadPath -ItemType Directory | Out-Null
}

# Install necessary Modules -if not present already
# install prerequisite NuGet
if (-Not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Install-PackageProvider -Name NuGet -Force
}

# install Evergreen powershell module
if (-Not (Get-Module -Name Evergreen -ErrorAction SilentlyContinue)) {
    Install-Module -Name Evergreen -Force
}

# update the Evergreen module - if already installed and not up to date
if (Get-Module -Name Evergreen -ErrorAction SilentlyContinue) {
    Update-Module -Name Evergreen -Force
}

# import evergreen module if it's not imported already
if (-Not (Get-Module -Name Evergreen -ErrorAction SilentlyContinue)) {
    Import-Module Evergreen
}

# update Evergreen app definitions
Update-Evergreen -Force

# get acrobat pro
$AcrobatReader = Get-EvergreenApp -Name AdobeAcrobatReaderDC | Where-Object { $_.Architecture -eq "x64" -and $_.Language -eq "MUI"}

# Get the File Name from the URI
$FileName = $AcrobatReader.uri.Split("/") | Select-Object -Last 1
Write-Host "Downloading $FileName from $($AcrobatReader.uri)"

# Get Version from File Name
$Version = $FileName.Substring(12, 10)
$Version = $Version.Insert(2, ".").Insert(6, ".")
Write-Host "Downloaded Adobe Acrobat DC version: $Version"

# Create a version folder if it does not exist
$VersionFolder = Join-Path $DownloadPath $Version
If (!(Test-Path -Path $VersionFolder)) {
    New-Item -Path $VersionFolder -ItemType Directory | Out-Null
}

$Installerpath = Join-Path $DownloadPath $FileName

# Download acrobat reader
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -UseBasicParsing $AcrobatReader.uri -OutFile $DownloadPath

# Unblock the downloaded file
$File = Join-Path $DownloadPath -ChildPath $FileName
Unblock-File -Path $File

# Expand the adobe reader exceutable
& $Installerpath -sfx_o"$VersionFolder" -sfx_ne -Wait

# Remove the zip file
Remove-Item -Path $File -Force

Write-Host "Adobe Acrobat Reader downloaded and expanded to $VersionFolder" -ForegroundColor Green