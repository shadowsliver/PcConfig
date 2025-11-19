<#
.SYNOPSIS
  Computer configuration script for ProCom

.DESCRIPTION
  Allows you to automate or configure various settings on a PC for ProCom deployments.

.INPUTS
  -Quick : Switch to run the script in quick mode, applying a predefined set of configurations without user interaction.

.OUTPUTS
  None.

.NOTES
  Version:        1.0
  Author:         Michel Michaux
  Creation Date:  23/10/2025
  Purpose/Change: Version 1.01

  This script needs to be run as Administrator. Executionpolicy should be set to RemoteSigned, Bypass or Unrestricted.

  If there are issues with the script, run it in terminal (openen PowerShell as Administrator).
  Use cd or Set-Location to navigate to the script folder
  .\Configure-ProCom.ps1

  Quick mode applies a predefined set of configurations without user interaction.

  If Winget is not working properly, use the debug mode to fix it.

  Openeing the windows store page and updating via the debug menu (option 3) usually works.

  Extra installations can be placed in the 'install' folder next to the script and configure the installations using the following files.
  .\Install\installations.csv
  EXAMPLE:
  /////////////////////////////////////
  software;file;arguments
  SOFTWARENAME;FILENAME.EXE;ARGUEMENTS
  SOFTWARENAME;FILENAME.EXE;ARGUEMENTS
  SOFTWARENAME;FILENAME.EXE;ARGUEMENTS
  /////////////////////////////////////
  Arguements can be empty

  Extra configurations can be placed in the 'config' folder next to the script.
  .\config\netdrive.csv
  EXAMPLE:
  /////////////////////////////////////
  Drive;Path
  N;\\server\data
  B;\\test\Folder1
  /////////////////////////////////////
  .\config\users.csv
  EXAMPLE:
  /////////////////////////////////////
  user;fullname;password
  JDoe;John Doe;Password123
  /////////////////////////////////////

.EXAMPLE
  Configure-ProCom.ps1
  Configure-ProCom.ps1 -Quick
#>


param (
  [switch] $Quick,
  [switch] $Advanced,
  [switch] $Debug
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

$AdminRequired = $true

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$logo = @"                                                                                    
                                                                                                              
      ..   .......     ...  ....    .......           .......       .......      .. .......  ........         
      +#*=########+.   .##+####=  .*########=.     .+########+.  .=########*.    ##########+:########-        
      +####+:..:+###=  .#####-.  =###=:..-*###:  .=###+:..:+#:  :###*-..:=###+.  ####+:.:+#####=::-###+.      
      +###.      .###: .###=.   +##+.      .###. :##*.          ###.      .*##-  ###:     :###.    .###.      
      +##=        :##+ .###     *##:        *#######=  -##########=         ###. ###.     .###.     *##.      
      +##*.      .=##= .###     +##-       .*##+++##*. .-=======###.       =##+  ###.     .###.     *##.      
      +###*:   ..*##+  .###     .###*.    -###+. .####:    .-.  =###-.   .+###.  ###.     .###.     *##.      
      +############-.  .###      .=##########:     =##########.  :##########=.   ###.     .###.     *##.      
      +##- :*##*:                   .=###*:           -####=.       :*###=.                                   
      +##-                                                                                                    
      +##-                                                                                                    
      +##-                                                                                                    
                                                                                                              
"@

# Hardcoded for protability, can be placed in a config file later
$winget_programs = @(
  "Google Chrome:Google.Chrome",
  "VLC Player:VideoLAN.VLC",
  "E-ID Middleware:BelgianGovernment.eIDmiddleware",
  "E-ID Viewer:BelgianGovernment.eIDViewer",
  "Adobe Acrobat Reader:Adobe.Acrobat.Reader.64-bit"
)

$office_lnk = 
"Word.lnk",
"Excel.lnk",
"PowerPoint.lnk",
"Outlook (classic).lnk"

# Create Office configuration XML
$OfficeXML = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365BusinessRetail">
      <Language ID="nl-nl" />
      <ExcludeApp ID="Publisher" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@

$knownBugs = @(
  "Disabling taskbar Widgets via registry is currently blocked by Windows security, need to find a workaround.",
  "Sometimes quick mode does not install all software if run through the batch shortcut.",
  "Some minor info displays may not show up correctly in certain Windows versions.",
  "If a windows update fails, the system may not reboot automatically. The update failing is a microsof/windows issue, not a script issue."
)

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Run {  
  While ( $true) {
    Write-Host "Config script - choose an option. Choose 0 to quit or press CTRL+C."
    Write-Host "1. Install basic software"
    Write-Host "2. Install Microsoft Office 365 NL"
    Write-Host "3. Update all software via Winget"
    Write-Host "4. Enable Windows updates and reboot"
    Write-Host "5. Set up basic machine configuration"
    Write-host "6. Adjust user performance profile settings"
    Write-Host "7. Disable password change on next login of current user"
    Write-Host "8. Create local admin user"
    Write-Host "9. Change device name"
    Write-Host "10. Rename local user account name & rename user folder (Other user only!)"
    Write-Host "11. Extra installs: Install folder executables, map network drives and create local users from config files"
    Write-Host "0. Exit"
    Write-Host ""
    Write-Host "Quick: Type 'quick' for the default configuration proccess"
    Write-Host "DEBUG: Type 'debug' to enter debug mode for Winget"
    Write-Host ""
    $choice = Read-Host "Choice"
    Write-Host ""

    functionPicker -choice $choice
  }
}

function functionPicker {
  param($choice)

  switch ($choice) {
    0 {
      Write-Host "Exiting..."
      exit
    }
    1 {
      ChoicePicker_Software_Install
    }
    2 {
      ChoicePicker_Office
    }
    3 {
      ChoicePicker_Update
    }
    4 {
      ChoicePicker_Windows_Update
    }
    5 {
      ChoicePicker_Basic_Config
    }
    6 {
      ChoicePicker_Adjust_User_Performance_Profile
    }
    7 {
      ChoicePicker_Current_User_No_pass
    }
    8 {
      ChoicePicker_User
    }
    9 {
      ChoicePicker_Change_Device_Name
    }
    10 {
      functionPicker_Rename_User
    }
    11 {
      ChoicePicker_Extra_installs
    }
    "wintool" {
      Open-Windows-Tool
    }
    "debug" {
      Debug
    }
    "quick" {
      Quick_config
    }
    default {
      Write-Host "Invalid choice. Please try again."
      Write-Host ""
    }
  }
}

function Quick_config {
  Write-Host "Disclaimer: Quick configuration will apply a predefined set of configurations without user interaction." -ForegroundColor Red
  Write-Host "Currently known bugs/issues with quick configuration:" -ForegroundColor Red
  foreach ($bug in $knownBugs) {
    Write-Host "- $bug" -ForegroundColor Red
  }
  Press_To_Continue -Message "Press any key to continue with quick configuration or CTRL+C to abort..."
  Write-Host ""

  Write-Host "Starting quick configuration..." -ForegroundColor Black -BackgroundColor White
  Write-Host ""

  Write-Host "Setting up basic machine configuration..." -ForegroundColor Yellow
  ChoicePicker_Basic_Config
  Write-Host "Basic machine configuration completed." -ForegroundColor Green
  Write-Host ""

  Write-Host "Adjusting user performance profile settings for best performance..." -ForegroundColor Yellow
  ChoicePicker_Adjust_User_Performance_Profile -quick
  Write-Host ""

  Write-Host "Installing basic software via Winget..." --ForegroundColor Yellow
  foreach ($program in $winget_programs) {
    $parts = $program -split ":"
    $name = $parts[0]
    $id = $parts[1]

    Write-Host "Installing $name..."
    winget install --id $id -e --accept-source-agreements --accept-package-agreements
  }
  Write-Host "Basic software installation completed." -ForegroundColor Green
  Write-Host ""

  Write-Host "Updating all installed software via Winget..." -ForegroundColor Yellow
  ChoicePicker_Update
  Write-Host "All software updates completed." -ForegroundColor Green
  Write-Host ""

  
  Write-Host "Running extra installs from configuration files..." -ForegroundColor Yellow
  ChoicePicker_Extra_installs
  Write-Host "Extra installs completed." -ForegroundColor Green

  Write-Host "Installing Microsoft Office 365 with Dutch configuration..." -ForegroundColor Yellow
  ChoicePicker_Office
  Write-Host "Microsoft Office 365 installation completed." -ForegroundColor Green
  Write-Host ""

  Write-Host "Applying windows updates..." -ForegroundColor Yellow
  ChoicePicker_Windows_Update
  Write-Host "Windows update process done. If there were any update errors the system wasn't rebooted" -ForegroundColor Green
  
  Write-Host "-----------------------------------------" -ForegroundColor DarkMagenta -BackgroundColor White
  Write-Host "Quick configuration completed. A reboot is recommended." -ForegroundColor Green -BackgroundColor White
  Write-Host ""
}

function ChoicePicker_Software_Install {
  Write-Host "Which software needs to be installed?"
  $i = 1
  foreach ($program in $winget_programs) {
    $parts = $program -split ":"
    $name = $parts[0]
    $id = $parts[1]

    Write-Host "$i. $name"
    $i++
  }
  Write-Host "0. Back to main menu"
  Write-Host "`n"
  $choice = Read-Host "Choice"

  if ($choice -eq 0) {
    return
  }
  elseif ($choice -ge 1 -and $choice -le $winget_programs.Count) {
    $selected_program = $winget_programs[$choice - 1]
    $parts = $selected_program -split ":"
    $name = $parts[0]
    $id = $parts[1]

    Write-Host "Installing $name..."
    winget install --id $id -e --accept-source-agreements --accept-package-agreements
  }
  else {
    Write-Host "Invalid choice. Please try again."
    ChoicePicker_Software_Install
  }
}

function ChoicePicker_User {
  param(
    [string]$user = "",
    [string]$userFull = "",
    [SecureString] $userPassword = ""
  )
  if ($user -eq "") {
    $user = Read-Host "Enter a valid Username for the new local admin account"
    $userFull = Read-Host "Enter the Full Name for the new local admin account (leave blank for no full name)"
    $userPassword = Read-Host "Enter a Password for the new local admin account (leave blank for no password)"
  }
  
  if ($userFull -eq "") {
    $userFull = $user
  }

  $securePassword = ConvertTo-SecureString $userPassword -AsPlainText -Force

  Write-Host "Creating local admin user '$user'...", $userfull, $userPassword
  # Create a local user with no password
  if ($userPassword -ne "") {    
    write-Host "With password"
    New-LocalUser -Name $user -Password $securePassword -FullName $userFull
    Set-LocalUser -Name $user -PasswordNeverExpires $true

  }
  else {
    write-Host "No password"
    New-LocalUser -Name $user -NoPassword -FullName $userFull
    Set-LocalUser -Name $user -PasswordNeverExpires $true

  }
  
  # Add the user to the Administrators group
  Add-LocalGroupMember -Group "Administrators" -Member $user

  Write-Host "Local admin user '$user' created successfully."
}

function ChoicePicker_Update {
  Write-Host "Updating all installed software via Winget..."
  winget upgrade --all --silent --accept-source-agreements --accept-package-agreements --include-unknown
  Write-Host "All software updates completed."
  Write-Host ""
}

function ChoicePicker_Office {
  Write-Host "Installing Microsoft Office 365 with Dutch configuration..."

  # Save XML to file
  $xmlPath = "$env:USERPROFILE\Downloads\office-nl.xml"
  $OfficeXML | Out-File -Encoding UTF8 -FilePath $xmlPath

  # Install Office Deployment Tool
  winget install --id Microsoft.OfficeDeploymentTool -e

  # Locate setup.exe (default install path)
  $setupPath = "C:\Program Files (x86)\OfficeDeploymentTool"

  # Run Office setup with Dutch config
  Start-Process -FilePath "$setupPath\setup.exe" -ArgumentList "/configure `"$xmlPath`"" -Wait

  Remove-Item $setupPath -Force -Recurse
  Remove-Item $xmlPath -Force -Recurse
  
  
  Write-Host "Creating desktop and start menu shortcuts for office..."
  # Create Desktop shortcuts
  foreach ($lnk in $office_lnk) {
    $source = [System.IO.Path]::Combine("C:\ProgramData\Microsoft\Windows\Start Menu\Programs", $lnk)
    $destination = [System.IO.Path]::Combine($env:USERPROFILE, "Desktop", $lnk)
    Copy-Item -Path $source -Destination $destination
  }
  
  Write-Host "Microsoft Office 365 installation completed."
  Write-Host ""
}

function ChoicePicker_Current_User_No_pass {
  Write-Host "Disabling password change upon next login for current user: $env:USERNAME"
  #Set-LocalUser -Name $env:USERNAME -Password (ConvertTo-SecureString "" -AsPlainText -Force) -PasswordNeverExpires $true
  net user $env:USERNAME /logonpasswordchg:no
  Write-Host "Password disabled for user '$env:USERNAME'."
}

function ChoicePicker_Change_Device_Name {
  $newName = Read-Host "Enter the new device name"
  if ($newName -ne "") {
    Rename-Computer -NewName $newName -Force -Restart:$false
    Write-Host "Device name changed to '$newName'. A restart is required for the change to take effect."
  }
  else {
    Write-Host "No device name entered. Returning to main menu."
    Write-Host ""
  }
  
}


function ChoicePicker_Basic_Config {
  
  Write-Host "Enabling Num-Lock on startup..." -ForegroundColor Green
  Set-ItemProperty -Path 'Registry::HKU\.DEFAULT\Control Panel\Keyboard' -Name "InitialKeyboardIndicators" -Value "2"
  

  Write-Host "Setting power plan to High Performance and adjusting monitor/standby timeouts..." -ForegroundColor Green
  # Set power plan to High Performance
  powercfg -setactive SCHEME_MIN
  
  # Set monitor timeout on AC power to 30 minutes
  powercfg /change monitor-timeout-ac 30

  # Set monitor timeout on battery (DC) power to never
  powercfg /change monitor-timeout-dc 0

  # Set standby timeout on AC power to never
  powercfg /change standby-timeout-ac 0

  # Set standby timeout on battery (DC) power to never
  powercfg /change standby-timeout-dc 0

  #Set Hard disk timeout on AC power to never
  #Set Hard disk timeout on battery (DC) power to never
  powercfg -change -disk-timeout-ac 0
  powercfg -change -disk-timeout-dc 0


  Write-Host "Disabling Hibernation and Fast Startup..." -ForegroundColor Green
  # Schakel Snel Opstarten uit
  #powercfg /hibernate off -> This also disables hibernation
  #HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power
  Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name 'HiberbootEnabled' -Value 0

  Write-Host "Enabling System Restore and creating a restore point..." -ForegroundColor Green
  # Enable System Restore if not already enabled
  # Enable system protection for C: drive
  Enable-ComputerRestore -Drive "C:\"

  Write-Host "Editing the taskbar" -ForegroundColor Green
  # Disable Search (0 = hidden, 1 = icon only, 2 = search box)
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 1

  # Disable Widgets ERROR: Currently blocked by windows, need to find a workaround
  #Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarDa" -Value 0

  # Disable Task View
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0

  Write-Host "Changing the user theme to Windows(light)"
  # Enable Light theme for system and apps
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 1
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 1

  $themePath = "C:\Windows\Resources\Themes\aero.theme"
  if (Test-Path $themePath) {
    Start-Process $themePath
  }
  else {
    Write-Host "Glow theme not found at $themePath"
  }

  # Restart Explorer to apply changes
  Stop-Process -Name explorer -Force
  Start-Process explorer

  Write-Host "Basic machine configuration completed." -BackgroundColor Green
  Write-Host ""
}

function Open-Windows-Tool {
  Invoke-WebRequest -useb https://christitus.com/win | Invoke-Expression
}

function ChoicePicker_Windows_Update {
  Write-Host "Running Windows Update to install all pending updates..." -ForegroundColor Green

  # 0. Zorg dat NuGet automatisch wordt geaccepteerd
  # Ensure TLS 1.2 is used (required for secure downloads)
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

  # Trust the PSGallery repository
  Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

  # Install NuGet provider silently
  $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
  if (-not $nugetProvider) {
    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -Confirm:$false
  }

  # 1. Installeer de PSWindowsUpdate-module (indien nodig)
  if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
  }

  # 2. Importeer de module
  Import-Module PSWindowsUpdate

  # 3. Voer alle beschikbare updates uit, inclusief optionele
  Get-WindowsUpdate -Install -AcceptAll -AutoReboot -Verbose
}

function Debug {
  $check = $true

  While ( $check -eq $true) {
    Write-Host "DEBUG MODE for application installer (Winget). Choose 0 to quit or press CTRL+C." -ForegroundColor Red -BackgroundColor Yellow
    Write-Host "1. Reregister Application Installer"
    Write-Host "2. Winget Update App Installer"
    Write-Host "3. See Application Installer via Microsoft Store (recommended first)"
    Write-Host "4. Reinstall Application Installer"
    Write-Host ""
    $choice = Read-Host "Choice"

    switch ($choice) {
      0 {
        Write-Host "Exiting debug mode..."
        $check = $false
      }
      1 {
        Write-Host "Reregistering Application Installer..."
        Get-AppxPackage Microsoft.DesktopAppInstaller | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" }
      }
      2 {
        Write-Host "Updating Application Installer via Winget..."
        winget upgrade Microsoft.AppInstaller
      }
      3 {
        Write-Host "See Application Installer via Microsoft Store..."
        Start-Process "ms-windows-store://pdp/?productid=9NBLGGH4NNS1"
      }
      4 {
        Invoke-WebRequest -Uri https://aka.ms/getwinget -OutFile .\WinGetSetup.exe
        Start-Process .\WinGetSetup.exe
      }
      default {
        Write-Host "Invalid choice. Please try again."
        Write-Host "`n"
      }
    }  
  }
}

function functionPicker_Rename_User {
  Write-Host "Listing all (enabled) local users:" -ForegroundColor Green
  get-LocalUser | Where-Object { $_.Enabled -eq $true } | Select-Object -Property Name
  
  
  $user = Read-Host "Enter the current username of the local user to rename, press 0 to cancel"

  if ($user -eq "0") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    return
  }

  $newUser = Read-Host "Enter the new username for the local user, press 0 to cancel"

  if ($newUser -eq "0") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    return
  }
  elseif ($user -eq "" -or $newUser -eq "") {
    Write-Host "Invalid input. Both current and new usernames must be provided." -ForegroundColor Red
    return
  }
  elseif ($user -eq $env:USERNAME) {
    Write-Host "You cannot rename the currently logged-in user. Please log in as a different user and try again." -ForegroundColor Red
    return
  }

  # Define paths
  $oldProfilePath = "C:\Users\$user"
  $newProfilePath = "C:\Users\$newUser"
  
  # 1. Rename the local user account
  Rename-LocalUser -Name $user -NewName $newUser

  # 2. Rename the user profile folder
  Rename-Item -Path $oldProfilePath -NewName $newProfilePath
  # 3. Get the SID of the renamed user
  $sid = (Get-LocalUser -Name $newUser | ForEach-Object {
      $user = $_
      $obj = New-Object System.Security.Principal.NTAccount($user.Name)
      $sid = $obj.Translate([System.Security.Principal.SecurityIdentifier])
      $sid.Value
    })

  # 4. Update the registry to point to the new profile path
  $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
  Set-ItemProperty -Path $regPath -Name "ProfileImagePath" -Value $newProfilePath

  Write-Host "User and profile folder renamed successfully. Please reboot the system." -ForegroundColor Green


}

#Do not use, have to properly research which versions is used for what and where
function ChoicePicker_Enable_AD_Tools {
  if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Host "Running in PowerShell Core"
    Install-Module -Name WindowsCompatibility
    Import-Module -Name WindowsCompatibility
    Import-WinModule -Name ActiveDirectory
  }
  elseif ($PSVersionTable.PSEdition -eq 'Desktop') {
    Write-Host "Running in Windows PowerShell"
    Write-Host "Installing Active Directory components..." -ForegroundColor Green
    # Install RSAT tools for Active Directory
    Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
  }
  else {
    Write-Host "Unknown PowerShell edition, please check manually." -ForegroundColor Red
  }
}

function ChoicePicker_Adjust_User_Performance_Profile {
  Param (
    [switch]$quick
  )
  $profileChoice = 2
  Write-Host "Adjusting user performance profile settings..." -ForegroundColor Green

  $check = $true
  if ( $quick -eq $true) {
    $check = $false
  }
  While ( $check -eq $true) {
    Write-host "Select the desired performance profile, press 0 to quit"
    Write-Host "1. Let windows choose what's best"
    Write-Host "2. Adjust for best appearance"
    Write-Host "3. Adjust for best performance"
    $profileChoice = Read-Host "Choice"
    if ($profileChoice -is [int] -and $profileChoice -in 1..3) {
      $check = $false
      <#
      - 0 = Let Windows choose what's best
      - 1 = Adjust for best appearance
      - 2 = Adjust for best performance
      #>
      $profileChoice = [int]$profileChoice - 1
    }
    elseif ($profileChoice -eq 0) {
      Write-Host "Operation cancelled by user." -ForegroundColor Yellow
      return
    }
    else {
      Write-Host "Invalid choice. Please try again." -ForegroundColor Red
    }
  }
  
  # Registry path for visual effects settings
  $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"

  # Set VisualFXSetting based on user choice
  Set-ItemProperty -Path $regPath -Name VisualFXSetting -Value $profileChoice

  # Apply changes by refreshing user settings
  RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters
  
  Write-Host "User performance profile adjusted. Please reboot for all changes to apply." -ForegroundColor Green
}

function ChoicePicker_Configure_IPv4 {
  #No error handling yet, use with caution
  Write-Host "Configuring IPv4 settings..." -ForegroundColor Green
  $interface = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1

  if ($null -eq $interface) {
    Write-Host "No network adapter found with an active connection." -ForegroundColor Red
    return
  }

  Write-Host "Network adapter "$interface.Name" found:" -ForegroundColor Green
  Write-Hoste "1. Configure a static IP address" -ForegroundColor Yellow
  Write-Host "2. Clear configuration (Dynamic)" -ForegroundColor Yellow
  $choice = Read-Host "Choose an option (1 or 2)"

  if ($choice -eq "2") {
    Write-Host "Clearing IPv4 configuration to dynamic (DHCP)..." -ForegroundColor Green
    Set-NetIPInterface -InterfaceAlias $interface.Name -Dhcp Enabled
    Remove-NetIPAddress -InterfaceAlias $interface.Name -AddressFamily IPv4 -Confirm:$false
    Write-Host "IPv4 configuration set to dynamic (DHCP)." -ForegroundColor Green
    return
  }
  elseif ($choice -eq 1) {
    $ipAddress = Read-Host "Enter the static IP address 0.0.0.0"
    $subnetMask = Read-Host "Enter the subnet mask"
    $defaultgateway = Read-Host "Enter the default gateway"
    $dnsServers = Read-Host "Enter the DNS servers (comma-separated if multiple)"

    Write-Host "Configuring static IPv4 settings..." -ForegroundColor Green
    New-NetIPAddress -InterfaceAlias $interface.Name -IPAddress $ipAddress -PrefixLength $subnetMask -DefaultGateway $defaultgateway
    $dnsServerArray = $dnsServers -split ","
    Set-DnsClientServerAddress -InterfaceAlias $interface.Name -ServerAddresses $dnsServerArray
    Write-Host "Static IPv4 configuration applied." -ForegroundColor Green
  }
  else {
    Write-host "Invalid choice. Skipping back to main menu" -ForegroundColor Red
    return
  }
}

#Automatic installation without a CSV
function ChoicePicker_Install_Install_Folder {
  Write-Host "Installing all files in the ./install/ folder..." -ForegroundColor Cyan
  #check if install folder exists
  $installPath = Join-Path -Path $PSScriptRoot -ChildPath "install"
  $check = Test-Path -Path $installPath
  if (-not $check) {
    Write-Host "Install folder not found. Please ensure the 'install' folder exists in the script directory." -ForegroundColor Red
    Write-Host ""
    return
  }

  $check = Get-ChildItem -Path $installPath -Filter *.exe
  if ($check.Count -eq 0) {
    Write-Host "No executable files found in the install folder." -ForegroundColor Red
    Write-Host ""
    return
  }

  $installPath = Join-Path -Path $PSScriptRoot -ChildPath "install"
  Write-Host "Install path: $installPath"
  $files = Get-ChildItem -Path $installPath -Filter *.exe
  foreach ($file in $files) {
    $filePath = $file.FullName
    Write-host "Installing: $filePath" -ForegroundColor Yellow
    Start-Process -FilePath $filePath -ArgumentList '/verysilent' -Wait
    Write-Host "Installed $($file.Name)" -ForegroundColor Green
  }

  Write-Host "Install done!" -ForegroundColor Green -BackgroundColor White
  Write-Host ""
  Write-Host ""
}

#Checks for the install/installations.csv
function ChoicePicker_Install_Install_Folder_Dyn {
  
  #check if install folder exists
  $installPath = Join-Path -Path $PSScriptRoot -ChildPath "install"
  $check = Test-Path -Path $installPath
  if (-not $check) {
    Write-Host "Install folder not found. Please ensure the 'install' folder exists in the script directory." -ForegroundColor Red
    Write-Host ""
    return
  }
  else {
    Write-Host "Installing all files in $installPath..." -ForegroundColor Cyan
  }

  $csvPath = Join-Path -Path $PSScriptRoot -ChildPath 'install\installations.csv'
  Write-Host "Reading CSV: $csvPath"

  $check = $true
  if (-not (Test-Path $csvPath)) {
    Write-Host "CSV not found: $csvPath"
    $check = $false
  }
  $rows = Import-Csv -Path $csvPath -Delimiter ";" -ErrorAction Stop

  if ($rows.Count -eq 0) {
    Write-Host "CSV is empty: $csvPath"
    $check = $false
  }

  if ($check -eq $true) {
    foreach ($row in $rows) {
      $fileName = $row.software
      $filePath = Join-Path -Path $installPath -ChildPath $row.file
      $fileArguement = $row.arguments
      if ($fileArguement -ne "") {
        Start-Process -FilePath $filePath -ArgumentList '/verysilent', $fileArguement -Wait
      }
      else {
        Start-Process -FilePath $filePath -ArgumentList '/verysilent' -Wait
      }
      Write-Host "Installed $FileName" -ForegroundColor Green  
    }  
  }

  Write-Host "Install done!" -ForegroundColor Green -BackgroundColor White
  Write-Host ""
  Write-Host ""
}

function ChoicePicker_Net_Stat_config {
  # Read ./config/netdrive.csv (relative to this script) and write out path + drive letter
  $csvPath = Join-Path -Path $PSScriptRoot -ChildPath 'config\netdrive.csv'
  Write-Host "Reading CSV: $csvPath"

  if (-not (Test-Path $csvPath)) {
    Write-Host "CSV not found: $csvPath"
    return
  }

  $rows = Import-Csv -Path $csvPath -Delimiter ";" -ErrorAction Stop

  if ($rows.Count -eq 0) {
    Write-Host "CSV is empty: $csvPath"
    return
  }

  foreach ($row in $rows) {
    Write-Host "Drive: $($row.Drive), Path: $($row.Path)"
    #New-PSDrive -Name $row.Drive -PSProvider "FileSystem" -Root $row.Path -Persist
    $driveLetter = "$($row.Drive):"
    $networkPath = $row.Path
    net use $driveLetter $networkPath /persistent:yes
  }
}

function ChoicePicker_Extra_installs {
  Write-Host "Running Installation folder executables..." -ForegroundColor Yellow
  ChoicePicker_Install_Install_Folder_Dyn
  Write-Host "Installation folder executables completed." -ForegroundColor Green
  Write-Host ""

  Write-Host  "Running Net Drive mapping from ./config/netdrive.csv..." -ForegroundColor Yellow
  ChoicePicker_Net_Stat_config
  Write-Host "Network drive mapping completed." -ForegroundColor Green
  Write-Host ""
  
  #Creating users
  Write-Host "Creating local admin users..." -ForegroundColor Yellow
  $csvPath = Join-Path -Path $PSScriptRoot -ChildPath 'config\users.csv'
  Write-Host "Reading CSV: $csvPath"

  $check = $true
  if (-not (Test-Path $csvPath)) {
    Write-Host "CSV not found: $csvPath"
    $check = $false
  }

  $rows = Import-Csv -Path $csvPath -Delimiter ";" -ErrorAction Stop

  if ($rows.Count -eq 0) {
    Write-Host "CSV is empty: $csvPath"
    $check = $false
  }

  if ($check -eq $true) {
    foreach ($row in $rows) {
      ChoicePicker_User -user $row.user -userFull $row.fullname -userPassword $row.password
    }  
  }
}

function Press_To_Continue {
  param(
    $Message = "Press any key to continue..."
  )
  Write-Host $message
  $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
Clear-Host
Write-Host $logo -foregroundColor DarkMagenta -BackgroundColor White
Write-Host "This script is for advanced users only. Use at your own risk!" -ForegroundColor Red
Write-Host ""

# Check if the script is running with administrative privileges
if ($AdminRequired -eq $true) {
  if (-not ([bool](New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Host "This script must be run as an administrator elevated window." -ForegroundColor Red
    Press_To_Continue -Message "Press any key to exit..."
    exit
  }
}

if ($Quick -eq $true) {
  winget source update
  Quick_config
}
elseif ($Advanced -eq $true) {
  Open-Windows-Tool
}
elseif ($Advanced -eq $true -and $Quick -eq $true) {
  Open-Windows-Tool
}
elseif ($Debug -eq $true) {
  Debug
}
else {
  Run
}
