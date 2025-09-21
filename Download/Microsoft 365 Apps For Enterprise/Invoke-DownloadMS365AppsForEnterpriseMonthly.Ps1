<#
.SYNOPSIS
    Downloads Micrsoft 365 Apps For Enterprise Monthly.
.DESCRIPTION
    This script downloads the latest version of Visual Studio Code and saves it to a specified directory.
.NOTES
    FileName:    Invoke-DownloadMS365AppsForEnterpriseMonthly.Ps1
    Author:      Dag Langfloen
    Curent Version: 1.0.1
    Version:     1.0.0 - Initial version
    Version:     1.0.1 - Changed folder structure to support Developer, and some error handling
    

#>

# Start-location to the script directory
$scriptStartDir = Get-Location

# Folders for Microsoft 365 Apps For Enterprise Monthly
$ApplicationFolderName = "Applications"
$Developer = "Microsoft Corporation"
$MS365AppsFolderName = "Microsoft 365 Apps For Enterprise"
$Channel = "Monthly Enterprise"
$SystemDrive = $env:SystemDrive
$ApplicationsFolder = Join-Path  $SystemDrive $ApplicationFolderName
$MS365AppsFolder = Join-Path $ApplicationsFolder -ChildPath $Developer | Join-Path -ChildPath $MS365AppsFolderName
$DownloadFolder = Join-Path $MS365AppsFolder $Channel

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

# Get the latest version of Microsoft 365 Apps For Enterprise Monthly from the XML feed
$xmlUrl = "https://docs.microsoft.com/en-us/officeupdates/monthly-enterprise-channel"
$page = Invoke-WebRequest -Uri $xmlUrl
$version = ($page.Content -split 'Version ')[1] -split '<' | Select-Object -First 1
Write-Output "Latest Office Monthly Enterprise Channel Version: $version"
# Split version into major, minor, build, revision
$versionParts = $version -split ': '
$Version = $versionParts[0]

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

# Find the lateest Office Deployment Tool
$ODTUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
$ODTPage = Invoke-WebRequest -Uri $ODTUrl
$ODTLink = ($ODTPage.Links | Where-Object { $_.href -like "*.exe" } | Select-Object -First 1).href

# Get the installer name from the url
$ODTInstaller = $ODTLink.Split('/')[-1]

# Offise Deployment Tool Folder
$ODTFolderName = "Office Deployment Tool"
$ODTFolderPath = Join-Path $MS365AppsFolder $ODTFolderName
$ODTPath = Join-Path $ODTFolderPath $ODTInstaller

# Create ODT folder if it doesn't exist
If (!(Test-Path $ODTFolderPath)) {
    try {
        New-Item -Path $ODTFolderPath -ItemType Directory | Out-Null -ErrorAction Stop    
    }
    catch {
        Write-Host "Error creating folder: $ODTFolderPath"
        $Error
    }
    
}   

# Check if ODT installer already exists and delete if it does
If (Test-Path $ODTPath) {
    try {
        Remove-Item $ODTPath -Force -ErrorAction SilentlyContinue    
    }
    catch {
        Write-Host "Error deleting file: $ODTPath" -ForegroundColor Red
        $Error
    }
    
}

# Download the Office Deployment Tool installer
try {
    Invoke-WebRequest -Uri $ODTLink -OutFile $ODTPath -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Office Deployment Tool from: $ODTLink" -ForegroundColor Red
    $Error
}

Write-Output "Downloaded Office Deployment Tool to: $ODTPath"

#unblock the ODT installer
Unblock-File -Path $ODTPath

# change to the ODT folder
Set-Location -Path $ODTFolderPath

# Extract the Office Deployment Tool ( requires admin rights)
try {
    Start-Process -FilePath $ODTInstaller -ArgumentList "/quiet", "/extract:.\" -Wait -ErrorAction Stop    
}
catch {
    Write-Host "Error extracting Office Deployment Tool: $ODTInstaller" -ForegroundColor Red
    $Error
}

Write-Output "Extracted Office Deployment Tool to: $ODTFolderPath"

# Copy the setup.exe from the ODT folder to the version folder
$ODTSetupPath = Join-Path $ODTFolderPath "setup.exe"
$DestinationSetupPath = Join-Path $VersionFolder "setup.exe"
try {
    Copy-Item -Path $ODTSetupPath -Destination $DestinationSetupPath -Force -ErrorAction Stop    
}
catch {
    Write-Host "Error copying setup.exe to: $DestinationSetupPath" -ForegroundColor Red
    $Error
}

Write-Output "Copied setup.exe to: $DestinationSetupPath"

# Create the configuration XML for downloading Microsoft 365 Apps For Enterprise Monthly
$ConfigXmlPath = Join-Path $VersionFolder "Configuration-Monthly.xml"
$ConfigXmlContent = @"
<Configuration ID="8d204ef3-4b4d-4979-9292-09519e6efa29">
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" MigrateArch="TRUE">
    <Product ID="O365ProPlusRetail">
      <Language ID="MatchOS" />
      <Language ID="MatchPreviousMSI" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="Publisher" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
$ConfigXmlContent | Out-File -FilePath $ConfigXmlPath -Encoding UTF8
Write-Output "Created configuration XML at: $ConfigXmlPath"

# Create the configuration xml for uninstalling Microsoft 365 Apps For Enterprise
$UninstallXmlPath = Join-Path $VersionFolder "Uninstall.xml"    
$UninstallXmlContent = @"
<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1"/>
</Configuration>
"@
$UninstallXmlContent | Out-File -FilePath $UninstallXmlPath -Encoding UTF8
Write-Output "Created uninstall XML at: $UninstallXmlPath"

# Change to the version folder
Set-Location -Path $VersionFolder

# Download Microsoft 365 Apps For Enterprise Monthly
try {
    Start-Process -FilePath setup.exe -ArgumentList "/download Configuration-Monthly.xml" -Wait -ErrorAction Stop    
}
catch {
    Write-Host "Error downloading Microsoft 365 Apps For Enterprise Monthly" -ForegroundColor Red
    $Error
}

Write-Output "Downloaded Microsoft 365 Apps For Enterprise Monthly to: $VersionFolder"

# The setup.exe, configuration XML and uninstall XML are now in the version folder
Write-Output "Microsoft 365 Apps For Enterprise Monthly setup files are located in: $VersionFolder"

# Return to the script start directory
Set-Location -Path $scriptStartDir