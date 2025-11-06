param (
  [switch] $Quick,
  [switch] $Advanced
)
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
  
.EXAMPLE
  Configure-ProCom.ps1
  Configure-ProCom.ps1 -Quick
#>

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

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Run {
  # Check if the script is running with administrative privileges
  if ($AdminRequired -eq $true) {
    if (-not ([bool](New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
      Write-Host "This script must be run as an administrator elevated window. Exiting..." -ForegroundColor Red
      exit
    }
  }
    
  :outerLoop while ($true) {
    :innerLoop While ( $true) {
      Write-Host "Config script - choose an option. Choose 0 to quit or press CTRL+C."
      Write-Host "1. Install basic software"
      Write-Host "2. Install Microsoft Office 365 NL"
      Write-Host "3. Create local admin user"
      Write-Host "4. Disable password change on next login of current user"
      Write-Host "5. Update all software via Winget"
      Write-Host "6. Change device name"
      Write-Host "7. Enable Windows updates and reboot"
      Write-Host "8. Set up basic machine configuration"
      Write-Host "9. Quick mode"
      Write-Host "0. Exit"
      Write-Host ""
      $choice = Read-Host "Choice"

      functionPicker -choice $choice
    }
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
      ChoicePicker_User
    }
    4 {
      ChoicePicker_Current_User_No_pass
    }
    5 {
      ChoicePicker_Update
    }
    6 {
      ChoicePicker_Change_Device_Name
    }
    7 {
      ChoicePicker_Windows_Update
    }
    8 {
      ChoicePicker_Basic_Config
    }
    9 {
      Quick_config
    }
    "wintool" {
      Open-Windows-Tool
    }
    "debug" {
      Debug
    }
    10 {
      Write-Host "Option 10 selected."
    }
    default {
      Write-Host "Invalid choice. Please try again."
      Write-Host "`n"
    }
  }
}

function Quick_config {
  Write-Host "Starting quick configuration..." -ForegroundColor Black backgroundColor White
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

  Write-Host "Installing Microsoft Office 365 with Dutch configuration..." -ForegroundColor Yellow
  ChoicePicker_Office
  Write-Host "Microsoft Office 365 installation completed." -ForegroundColor Green
  Write-Host ""

  Write-Host "Disabling password change on next login for current user: $env:USERNAME" -ForegroundColor Yellow
  ChoicePicker_Current_User_No_pass
  Write-Host "Password disabled for user '$env:USERNAME'." -ForegroundColor Green
  Write-Host ""

  Write-Host "Updating all installed software via Winget..." -ForegroundColor Yellow
  ChoicePicker_Update
  Write-Host "All software updates completed." -ForegroundColor Green
  Write-Host ""

  Write-Host "Setting up basic machine configuration..." -ForegroundColor Yellow
  ChoicePicker_Basic_Config
  Write-Host "Basic machine configuration completed." -ForegroundColor Green
  Write-Host ""

  Write-Host "Quick configuration completed. A reboot is recommended.'n'n" -ForegroundColor Green backgroundColor White
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
  $user = Read-Host "Enter a valid Username for the new local admin account"
  $userFull = Read-Host "Enter the Full Name for the new local admin account (leave blank for no full name)"
  $userPassword = Read-Host "Enter a Password for the new local admin account (leave blank for no password)"
  if ($userFull -eq "") {
    $userFull = $user
  }
  # Create a local user with no password
  if ($userPassword -ne "") {
    $securePassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
    New-LocalUser -Name $user -Password $securePassword -FullName $userFull -PasswordNeverExpires $true -UserMayNotChangePassword $false
  }
  else {
    New-LocalUser -Name $user -NoPassword -FullName $userFull -PasswordNeverExpires $true -UserMayNotChangePassword $false
  }
  
  # Add the user to the Administrators group
  Add-LocalGroupMember -Group "Administrators" -Member $user

  Write-Host "Local admin user '$user' created successfully.'n'n"
}

function ChoicePicker_Update {
  Write-Host "Updating all installed software via Winget..."
  winget upgrade --all --silent --accept-source-agreements --accept-package-agreements --include-unknown
  Write-Host "All software updates completed.'n'n"
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
  
  Write-Host "Microsoft Office 365 installation completed.'n'n"
}

function ChoicePicker_Current_User_No_pass {
  Write-Host "Disabling password change upon next login for current user: $env:USERNAME"
  #Set-LocalUser -Name $env:USERNAME -Password (ConvertTo-SecureString "" -AsPlainText -Force) -PasswordNeverExpires $true
  net user $env:USERNAME /logonpasswordchg:no
  Write-Host "Password disabled for user '$env:USERNAME'.'n'n"
}

function ChoicePicker_Change_Device_Name {
  $newName = Read-Host "Enter the new device name"
  if ($newName -ne "") {
    Rename-Computer -NewName $newName -Force -Restart:$false
    Write-Host "Device name changed to '$newName'. A restart is required for the change to take effect.'n'n"
  }
  else {
    Write-Host "No device name entered. Returning to main menu.'n'n"
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
  powercfg /hibernate off

  Write-Host "Enabling System Restore and creating a restore point..." -ForegroundColor Green
  # Enable System Restore if not already enabled
  # Enable system protection for C: drive
  Enable-ComputerRestore -Drive "C:\"


  Write-Host "Basic machine configuration completed.'n'n" -BackgroundColor Green
}

function Open-Windows-Tool {
  Invoke-WebRequest -useb https://christitus.com/win | Invoke-Expression
}

function ChoicePicker_Windows_Update {
  Write-Host "Running Windows Update to install all pending updates..." -ForegroundColor Green
  # 1. Install the module (if not already installed)
  Install-Module -Name PSWindowsUpdate -Force

  # 2. Import the module
  Import-Module PSWindowsUpdate

  # 3. Run all available updates, including optional ones
  Get-WindowsUpdate -Install -AcceptAll -AutoReboot
}

function Debug {
  $check = $true

  While ( $check -eq $true) {
    Write-Host "DEBUG MODE for application installer (Winget). Choose 0 to quit or press CTRL+C." -ForegroundColor Red -BackgroundColor Yellow
    Write-Host "1. Reregister Application Installer"
    Write-Host "2. Winget Update App Installer"
    Write-Host "3. See Application Installer via Microsoft Store..."
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

#-----------------------------------------------------------[Execution]------------------------------------------------------------
winget source update
Clear-Host
Write-Host $logo -foregroundColor DarkMagenta -BackgroundColor White

# Check if the script is running with administrative privileges
if ($AdminRequired -eq $true) {
  if (-not ([bool](New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Host "This script must be run as an administrator elevated window." -ForegroundColor Red
    Write-Host "Press any key to terminate the script..."
    $x = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
  }
}

if ($Quick -eq $true) {
  Quick_config
}
elseif ($Advanced -eq $true) {
  Open-Windows-Tool
}
elseif ($Advanced -eq $true -and $Quick -eq $true) {
  Open-Windows-Tool
}
else {
  Run
}




