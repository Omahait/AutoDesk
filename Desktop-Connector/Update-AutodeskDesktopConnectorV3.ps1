<# 

############   Update Autodesk Desktop Connector Written by InfiNet Solutions   ############00

.synopsis
    This script will update Autodesk Desktop Connector to the latest version avaliable from Autodesk.

#>

# Function to update Autodesk Desktop Connector
function Update-AutoDeskDesktopConnector {
    param (
    )
    $InstallPath = "C:\Autodesk\"

  # Look for Autodesk Desktop Connector Path and create if needed
    if (!(Test-Path $InstallPath)) {
      New-Item -ItemType Directory -Path $InstallPath
  }
  if (Test-Path "$InstallPath\AutodeskDesktopConnectorSetup.exe") {
      Remove-Item -ItemType directory -Path "$InstallPath" -Recurse
  }

  # Delete old installer file if it exists
  $oldInstaller = "$InstallPath\DesktopConnector-64.exe"
  if (Test-Path $oldInstaller) {
    Remove-Item $oldInstaller
  }
  else {
    Write-Host "No old installer found. Continuing..."
  }
# Kill Autodesk Desktop Connector Service if running
$serviceName = "DesktopConnector.Applications.Tray"
$service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    
#check if the service exists and is running
Stop-Process -Name $serviceName -Force -ErrorAction SilentlyContinue

  # Check for older versions of Autodesk Desktop Connector
    $oldFolders = Get-ChildItem -Directory -Path $InstallPath | Where-Object {$_.Name -like "Autodesk_Desktop_Connector_*"} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($oldFolders -ne $null) {
      $OldSetupPath = Join-Path -Path $oldFolders.FullName -ChildPath "setup.exe"

        # Check if the setup.exe exists inside the folder and start uninstallation
        if (Test-Path $OldSetupPath -PathType Leaf) {
          $uninstallArgs = "-i uninstall --silent"
          Write-Host "Older version of Autodesk Desktop Connector found. Trying Uninstall..."
          Start-Process -FilePath $OldSetupPath -ArgumentList $uninstallArgs -Wait
        }
        else {
          Write-Host "setup.exe not found in the latest folder. Trying to install..."
        }
              
    }
    else {
      Write-Host "No older versions of Autodesk Desktop Connector found. Continuing..."
    }

  # Download Variables
  write-host "Downloading Autodesk Desktop Connector..."
  New-Item -ItemType Directory -Path "$installPath" -ErrorAction SilentlyContinue
  $ExeURI = "https://www.autodesk.com/adsk-connect-64"
  $DestinationFolder = "$InstallPath\DesktopConnector-64.exe"

  # Download installer
  $ProgressPreference = 'SilentlyContinue'
  Invoke-WebRequest -Uri $ExeURI -OutFile $DestinationFolder 

  # Error Handeling for failed download
  if (!(Test-Path "$DestinationFolder")) {
      Write-Host "Failed to download Autodesk Desktop Connector"
      exit 1
  }

  # Extract installer
    # Variables
  cd $InstallPath
  write-host "Extracting Autodesk Desktop Connector..."
  $setupExtractPackage = "$InstallPath\DesktopConnector-64.exe"
  $setupExtractArgs = "-suppresslaunch -d $InstallPath"
  
    # Start the extraction process
  Start-Process -FilePath $setupExtractPackage -ArgumentList $setupExtractArgs -Wait

  # Find New Installer
  $newestFolder = Get-ChildItem -Directory -Path $InstallPath | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($newestFolder -ne $null) {
    $setupPath = Join-Path -Path $newestFolder.FullName -ChildPath "setup.exe"

  # Check if the setup.exe exists inside the folder and start installation
  if (Test-Path $setupPath -PathType Leaf) {
  
    # Start the setup Process

    # Install Autodesk Desktop Connector
    $installArgs = "-i install --silent"
    Start-Process -FilePath $setupPath -ArgumentList $installArgs
    } else {
        Write-Host "setup.exe not found in the latest folder."
        exit 1
    }

  } 
  else {
    Write-Host "No subfolders found in $installerFolderPath."
    exit 1
  }
  
  # Restart Explorer if not running
  $processName = "explorer"
  if (!(Get-Process -Name $processName -ErrorAction SilentlyContinue)) {
    Start-Process -FilePath $processName
  }

Write-Host "Autodesk Desktop Connector setup finishing in the background. Please wait for the installation to finish."

# End of function
exit 0
}
 
# Run Function
Update-AutoDeskDesktopConnector

# End of script
