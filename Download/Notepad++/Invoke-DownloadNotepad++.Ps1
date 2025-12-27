<#
.SYNOPSIS
    Downloads Notepad++.
.DESCRIPTION
    This script downloads the latest version of Notepad++ and saves it to a specified Directory
.NOTES
    FileName:    Invoke-DownloadNotepad++.Ps1
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
$Developer = "Notepad++ Team"
$ApplicationName = "Notepad++"
$SystemDrive = $env:SystemDrive

# Download the latest Notepad++ installer (x64)
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

# Get latest release info from GitHub API
$apiUrl = "https://api.github.com/repos/notepad-plus-plus/notepad-plus-plus/releases/latest"
$release = Invoke-RestMethod -Uri $apiUrl

# Find the x64 installer asset
$asset = $release.assets | Where-Object { $_.name -like "*x64*.exe" } | Select-Object -First 1
$downloadUrl = $asset.browser_download_url

# Output the latest version tag
$latestVersion = $release.tag_name
Write-Host "Latest Notepad++ version: $latestVersion"

# Remove the 'v' prefix if present
if ($latestVersion -like "v*") {
    $latestVersion = $latestVersion.Substring(1)
}

# Create a folder with the version number
$VersionFolder = Join-Path $DownloadFolder $latestVersion   
If (!(Test-Path $VersionFolder)) {
    try {
        New-Item -Path $DownloadFolder -Name $latestVersion -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $VersionFolder"
        $Error
    }
    
}

# Get the installer name
$InstallerName = "npp.$latestVersion.Installer.x64.exe"

# Download path
$DownloadPath = Join-Path $VersionFolder  $InstallerName

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

# Download the installer
try {
    Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -OutFile $DownloadPath -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Notepad++ from: $downloadUrl" -ForegroundColor Red
    $Error
}
Write-Host "Downloaded Notepad++ installer to: $DownloadPath"