# PcConfig

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