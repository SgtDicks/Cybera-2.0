$Id: WinXP_readme.txt 18 2006-01-04 09:48:40Z sam $
________________________________________________________________________________
Cybera Workstation & Windows XP
________________________________________________________________________________
I) WinXP Home - WinXP Pro:
In it's present state, Cybera Workstation cannot work properly with WinXP Home. 
This is because of two limitations of that OS:
1) You cannot remotely access the registry (the remote registry access service 
   is not installed).
2) You cannot deactivate the option: "Simple file sharing" in "Folder Options"
   This means that a computer on the net has to connect to the computer with 
   only guest privileges. 
These two limitations mean that the server cannot connect to the workstation 
registry to lock down user rights. In a possible future release (say 2.x) I 
will add a true NT service on the workstation that should fix this but for now, 
Cybera is only truly compatible with WinXP pro and win2000.
________________________________________________________________________________
II) WinXP Pro configuration:
1) WinXP Pro also has "Simple file sharing" activated by default. You will have 
   to deactivate this on the workstations in order for user rights lockdown to 
   work.
2) In order to forbid logoff and shutdown you have to:
   a) Deactivate the WinXP Welcome screen (the screen where you click on the 
      user you want to login as). The installer package automatically does this 
      for you.
   b) Set the start menu to classical (if not the logoff button STILL is present
      even when it has been expressly forbidden in the registry). Cybera will do
      this for you as part of user rights lockdown (Forbid computer shutdown).
