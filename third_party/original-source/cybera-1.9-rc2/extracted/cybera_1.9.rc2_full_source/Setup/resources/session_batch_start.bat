@echo off
rem  ___________________________________________________________________________
rem |
rem | description: Session Start batch file.
rem |              This Bacth file will be run at the start of any session
rem |
rem | parameters:  %1 - Client login or card number. For self-service sessions
rem |                   this parameter defaults to %3
rem |              %2 - Server computer name
rem |              %3 - Workstation computer name
rem |
rem | Example usage: mount a network drive on Z: that points to a folder named
rem | after the client login, card number or workstation computer name in the 
rem | case of self service sessions. This example supposes that the server
rem | shares a folder called 'cli_files' that the workstations can access and
rem | write to.
rem |___________________________________________________________________________
rem BEGIN EXAMPLE
rem net use z: /delete /y
rem net use z: \\%2\cli_files
rem mkdir z:\%1
rem net use z: /delete /y
rem net use z: \\%2\cli_files\%1
rem END EXAMPLE

rem Add your commands here
