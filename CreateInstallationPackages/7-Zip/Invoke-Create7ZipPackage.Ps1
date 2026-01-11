<#
.SYNOPSIS
    Creates installation packages for 7-Zip.
.DESCRIPTION
    This script creates installation packages for the latest version of 7-Zip. Based on the downloaded installer.
.NOTES
    FileName:    Invoke-Create7ZipPackage.Ps1
    Author:      Dag Langfloen
    Current version: 1.2
    Version History:
        1.0 - Initial version
        1.1 - Added support for both MSI and EXE installers
        1.2 - Changed versioning og the script
#>

# Folders
$ApplicationFolderName = "Applications"
$Purpose = "InstallationPackages"
$ApplicationName = "7-Zip"
$SystemDrive = $env:SystemDrive

# Icon file path
$IconFile = "C:\Applications\Icons\7-Zip.png"

# Deplyment toolkit variables
# Application Info
$AppName = "7-Zip"
$AppVendor = "7-Zip.org"
$AppProcessesToClose = "AppProcessesToClose = @(@{ Name = '7zFM'; Description = '7-Zip FileManager'},@{ Name = '7zG'; Description = '7-Zip GUI'},@{ Name = '7z.exe'; Description = '7-Zip Console'})"
$AppScriptDate = Get-Date -Format "yyyy.MM.dd"
$AppScriptAuthor = ((Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\SessionData\" | Get-ItemProperty | Where-Object LoggedOnUser -like "$(whoami)" | select -First 1)).LoggedOnDisplayName
$RequireAdmin = $true

# Installation arguments for MSI
$InstallArguments = "/qn /norestart"   
$UnInstallArguments = "/qn /norestart"

# Processing the latest downloaded 7-Zip installer
# Get the latest downloaded version of 7-Zip
$DownloadFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath "Downloads" | Join-Path -ChildPath $ApplicationName
$LatestVersion = (Get-ChildItem $DownloadFolder | Sort-Object Name -Descending | Select-Object -First 1).Name
Write-Host "Latest downloaded 7-Zip version: $LatestVersion"

# ShowADTInstallationWelcome 
$OldShowADTInstallationWelcome = "Show-ADTInstallationProgress"
$ShowADTInstallationWelcome = "Processing $AppName $LatestVersion"
$NewShowADTInstallationWelcome = "Show-ADTInstallationProgress -StatusMessage '$ShowADTInstallationWelcome'" 

# ShowADTInstallationPrompt
$OldShowADTInstallationPrompt = "Show-ADTInstallationPrompt -Message 'You can customize text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait"
$ShowADTInstallationPrompt = "$AppName $LatestVersion was successfully installed."
$NewShowADTInstallationPrompt = "Show-ADTInstallationPrompt -Message '$ShowADTInstallationPrompt' -ButtonRightText 'OK' -Icon Information -NoWait"

# Old AppprocessesToClose string
$OldAppProcessesToClose = "AppProcessesToClose = @()  # Example: @('excel', @{ Name = 'winword'; Description = 'Microsoft Word' })"

# Get the filename of the downloaded installer
$DownloadedInstaller = Get-ChildItem -Path (Join-Path $DownloadFolder $LatestVersion) -Filter "*.msi" | Select-Object -First 1
If ($null -eq $DownloadedInstaller) {
    Write-Host "No downloaded installer found for version $LatestVersion in $DownloadFolder"
    Exit 1
}
$InstallerName = $DownloadedInstaller.Name
Write-Host "Found downloaded installer: $InstallerName"

# Original installation steps

# PreInstallation step
$OriginalPreInstallationStep = "## <Perform Pre-Installation tasks here>"

# Installation step
$OriginalInstallationStep = "## <Perform Installation tasks here>"

# PostInstallation step
$OriginalPostInstallationStep = "## <Perform Post-Installation tasks here>"

# PreUninstallation step
$OriginalPreUnInstallationStep = "## <Perform Pre-Uninstallation tasks here>"

# Uninstallation step
$OriginalUnInstallationStep = "## <Perform Uninstallation tasks here>"

# If the installer has file extension .msi, set the installer type to MSI
If ($InstallerName -like "*.msi") {
    
    # PreInstallation step
    $NewPreInstallationStep = "## Pre-Installation tasks" 

    # Installation step
    $NewInstallationStep = "## Installation tasks 
    Start-ADTMsiProcess -Action Install -FilePath '$InstallerName' -ArgumentList '$InstallArguments'"
    
    # PostInstallation step
    $NewPostInstallationStep = "## Post-Installation tasks"

    # PreUninstallation step
    $NewPreUnInstallationStep = "## Pre-Uninstallation tasks"

    # Uninstallation steps
    $NewUnInstallationStep = "## Uninstallation tasks 
   Start-ADTMsiProcess -Action Uninstall -FilePath '$InstallerName' -ArgumentList '$UnInstallArguments'"

    # PostUninstallation step
    $NewPostUnInstallationStep = "## Post-Uninstallation tasks"
}
Elseif ($InstallerName -like "*.exe") {
    # PreInstallation step
    $NewPreInstallationStep = "## Pre-Installation tasks" 

    # Installation step
    $NewInstallationStep = "## Installation tasks 
    Start-ADTProcess -FilePath '$InstallerName' -ArgumentList '$InstallArguments'"
    
    # PostInstallation step
    $NewPostInstallationStep = "## Post-Installation tasks"

    # PreUninstallation step
    $NewPreUnInstallationStep = "## Pre-Uninstallation tasks"

    # Uninstallation steps
    $NewUnInstallationStep = "## Uninstallation tasks 
   Start-ADTProcess -FilePath '$InstallerName' -ArgumentList '$UnInstallArguments'"

    # PostUninstallation step
    $NewPostUnInstallationStep = "## Post-Uninstallation tasks"
}



# Downloaded installer path
$InstallerPath = Join-Path -Path $DownloadFolder -ChildPath $LatestVersion | Join-Path -ChildPath $DownloadedInstaller.Name
Write-Host "Using installer: $InstallerPath"

# Create Installation Package folder
$InstallationPackageFolder = Join-Path  $SystemDrive -ChildPath $ApplicationFolderName | Join-Path -ChildPath $Purpose | Join-Path -ChildPath $ApplicationName | Join-Path -ChildPath $LatestVersion
If (!(Test-Path $InstallationPackageFolder)) {
    try {
        New-Item -Path $InstallationPackageFolder -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $InstallationPackageFolder"
        $Error
    }
    
}
else {
    Write-Host "Installation package folder already exists: $InstallationPackageFolder"
}

# Processing the App Deployment Toolkit
# Powershell App Deployment Toolkit Location
$PSADTLocation = "C:\Applications\Downloads\PSAppDeployToolkit"

# Get the latest version of PSADT
$PSADTVersion = (Get-ChildItem "C:\Applications\Downloads\PSAppDeployToolkit" | Sort-Object Name -Descending | Select-Object -First 1).Name
$PSADTPath = Join-Path -Path $PSADTLocation -ChildPath $PSADTVersion

# Copy the files in $psadtPath to the InstallationPackageFolder
Copy-Item -Path (Join-Path $PSADTPath "*") -Destination $InstallationPackageFolder -Recurse -Force
Write-Host "Copied PSADT files from $PSADTPath to $InstallationPackageFolder"

# Remove the .txt files in the Files folder and the supportfiles folder
$FilesFolder = Join-Path $InstallationPackageFolder "Files"
$SupportFilesFolder = Join-Path $InstallationPackageFolder "SupportFiles"
Remove-Item -Path (Join-Path $FilesFolder "*.txt") -Force
Remove-Item -Path (Join-Path $SupportFilesFolder "*.txt") -Force
Write-Host "Removed .txt files from Files and SupportFiles folders"

# Rename the AppIcon.png file in the Assets folder
$AssetsFolder = Join-Path $InstallationPackageFolder "Assets"
Rename-Item -Path (Join-Path $AssetsFolder "AppIcon.png") -NewName "Old_AppIcon.png"
Write-Host "Renamed existing AppIcon.png to Old_AppIcon.png"

# Copy the icon file to the Assets folder
Copy-Item -Path $IconFile -Destination (Join-Path $AssetsFolder "AppIcon.png") -Force
Write-Host "Copied icon file to Assets folder"

# Copy the installer to the Files folder
Copy-Item -Path $InstallerPath -Destination (Join-Path $FilesFolder $DownloadedInstaller.Name) -Force
Write-Host "Copied installer to Files folder"

# Get the app architecture from the filename    
If ($DownloadedInstaller.Name -match "x64") {
    $AppArchitecture = "x64"
} elseif ($DownloadedInstaller.Name -match "x86") {
    $AppArchitecture = "x86"
} else {
    $AppArchitecture = "Unknown"
}
Write-Host "Application Architecture: $AppArchitecture"

# Update Deploy-Application.ps1 script
$DeployScriptPath = Join-Path $InstallationPackageFolder "Invoke-AppDeployToolkit.ps1"

# Read the content of the Deploy-Application.ps1 script
$DeployScriptContent = Get-Content -Path $DeployScriptPath

# Replace ADTSession variables
$DeployScriptContent = $DeployScriptContent -replace "AppVendor = ''", "AppVendor = '$AppVendor'"
$DeployScriptContent = $DeployScriptContent -replace "AppName = ''", "AppName = '$AppName'"
$DeployScriptContent = $DeployScriptContent -replace "AppVersion = ''", "AppVersion = '$LatestVersion'"
$DeployScriptContent = $DeployScriptContent -replace "AppArch = ''", "AppArch = '$AppArchitecture'"
$DeployScriptContent = $DeployScriptContent -replace "AppProcessesToClose = @()  # Example: @('excel', @{ Name = 'winword'; Description = 'Microsoft Word' })", "AppProcessesToClose = $AppProcessesToClose"
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OldAppProcessesToClose), $AppProcessesToClose
$DeployScriptContent = $DeployScriptContent -replace "AppScriptDate = '2025-10-21'", "AppScriptDate = '$AppScriptDate'"
$DeployScriptContent = $DeployScriptContent -replace "AppScriptAuthor = '<author name>'", "AppScriptAuthor = '$AppScriptAuthor'"
$DeployScriptContent = $DeployScriptContent -replace "RequireAdmin = $true", "RequireAdmin = $RequireAdmin"

# InstallTitles
$DeployScriptContent = $DeployScriptContent -replace "InstallName = '.*?'", "InstallName = '$AppName $LatestVersion'"
$DeployScriptContent = $DeployScriptContent -replace "InstallTitle = '.*?'", "InstallTitle = 'Processing $AppName $LatestVersion'"

# ShowADTInstallationWelcome
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OldShowADTInstallationWelcome), $NewShowADTInstallationWelcome

# ShowADTInstallationPrompt
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OldShowADTInstallationPrompt), $NewShowADTInstallationPrompt

# Set pre-installation step
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OriginalPreInstallationStep), $NewPreInstallationStep

# Set installation step
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OriginalInstallationStep), $NewInstallationStep

# Set post-installation step
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OriginalPostInstallationStep), $NewPostInstallationStep

# Set pre-uninstallation step
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OriginalPreUnInstallationStep), $NewPreUnInstallationStep

# Set uninstallation step
$DeployScriptContent = $DeployScriptContent -replace [regex]::Escape($OriginalUnInstallationStep), $NewUnInstallationStep

# Write the updated content back to the file
Set-Content -Path $DeployScriptPath -Value $DeployScriptContent -Encoding UTF8

Write-Host "Updated Invoke-AppDeployToolkit.ps1 script with application details" -ForegroundColor Green