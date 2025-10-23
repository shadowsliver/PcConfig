<#
.SYNOPSIS
  <Overview of script>

.DESCRIPTION
  <Brief description of script>

.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>

.INPUTS
  <Inputs if any, otherwise state None>

.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>

.NOTES
  Version:        1.0
  Author:         Michel Michaux
  Creation Date:  23/10/2025
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------


#-----------------------------------------------------------[Functions]------------------------------------------------------------


$logo = @"
INSERT PROCOM LOGO HERE
"@
Write-Host $logo -ForegroundColor Green

function functionPicker {
    param($choice)

    switch ($choice) {
        0 {
            Write-Host "Exiting..."
            exit
        }
        1 {
            Write-Host "Choice 1"
        }
        2 {
            Write-Host "Choice 2"
        }
        3 {
            Write-Host "Choice 3"
        }
        4 {
            Write-Host "Choice 4"
        }
        default {
            Write-Host "Invalid choice. Please try again."
            functionPicker -choice $choice
        }
    }
}


# Check if the script is running with administrative privileges

if (-not ([bool](New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Host "This script must be run as an administrator. Exiting..."
    exit
}

function Run {
    :outerLoop while ($true) {
        Clear-Host
        :innerLoop While ( $true) {
            Write-Host "Config script - choose an option. Choose 0 to quit.`n`n"
            Write-Host "1. WIP"
            Write-Host "2. WIP"
            Write-Host "3. WIP"
            Write-Host "4. WIP"
            Write-Host "0. Exit"
            Write-Host "`n"
            $choice = Read-Host "Choice"

            functionPicker -choice $choice
        }
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

:outerLoop while ($true) {
    :innerLoop While ( $true) {
        Write-Host "DZB Command script - choose an option. Choose 0 to quit.`n`n"
        Write-Host "1. Add client(s)"
        Write-Host "2. Remove client"
        Write-Host "3. Execute command"
        Write-Host "4. Display current selected clients"
        Write-Host "0. Exit"
        Write-Host "`n"
        $choice = Read-Host "Choice"

        functionPicker -choice $choice
    }
}