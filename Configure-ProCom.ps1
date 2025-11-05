param (
  [switch] $Quick
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
  Purpose/Change: Version 1.0
  
.EXAMPLE
  Configure-ProCom.ps1
  Configure-ProCom.ps1 -Quick
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
$AdminRequired = $false

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

$winget_programs = @(
  "Google Chrome:Google.Chrome",
  "VLC Player:VideoLAN.VLC",
  "E-ID Middleware:BelgianGovernment.eIDmiddleware",
  "E-ID Viewer:BelgianGovernment.eIDViewer"
)

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
      Write-Host "7. Reboot device"
      Write-Host "8. Quick mode"
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
      Write-Host "Rebooting device..."
      Restart-Computer
    }
    8 {
      Quick_config
    }
    default {
      Write-Host "Invalid choice. Please try again."
      Write-Host "`n"
    }
  }
}

function Quick_config {

  if ($AdminRequired -eq $true) {
    if (-not ([bool](New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
      Write-Host "This script must be run as an administrator elevated window. Exiting..." -ForegroundColor Red
      exit
    }
  }

  Write-Host "Installing basic software via Winget..."
  foreach ($program in $winget_programs){
    $parts = $program -split ":"
    $name = $parts[0]
    $id = $parts[1]

    Write-Host "Installing $name..."
    winget install --id $id -e --accept-source-agreements --accept-package-agreements
  }
  Write-Host "Basic software installation completed.'n'n"

  Write-Host "Installing Microsoft Office 365 with Dutch configuration..."
  ChoicePicker_Office
  Write-Host "Microsoft Office 365 installation completed.'n'n"

  Write-Host "Disabling password change on next login for current user: $env:USERNAME"
  ChoicePicker_Current_User_No_pass
  Write-Host "Password disabled for user '$env:USERNAME'.'n'n"

  Write-Host "Updating all installed software via Winget..."
  ChoicePicker_Update
  Write-Host "All software updates completed.'n'n"

  Write-Host "Quick configuration completed. A reboot is recommended.'n'n"
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

#-----------------------------------------------------------[Execution]------------------------------------------------------------
winget source update
Clear-Host
Write-Host $logo -foregroundColor DarkMagenta -BackgroundColor White

if ($Quick -eq $true) {
  Quick_config
}
else {
  Run
}