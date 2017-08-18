Deploy-LAPS
=================
![Powershell v4+](https://img.shields.io/badge/Powershell-v4-blue.svg) [![License: GPL v3](https://img.shields.io/badge/License-GPL%20v3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)  [![contributions welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg?style=flat)](https://github.com/sup3rlativ3/deploy-laps/issues)

This script and accompanying files will allow system administrators to automatically deploy Microsoft Local Administrator Password Solution (LAPS). Please visit https://technet.microsoft.com/en-us/mt227395.aspx for more information about LAPS.



 Parameters
 -------------- 
 **-ComputerOU**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
  	This is the OU that LAPS will assign permissions to update it's own AD object to write the password when changed.
  	If this is not set it will pull the default computer OU from AD.  
  
 **-OrgUnitRead**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
  	Sets the OU that has read access to the passwords stored in the computer AD object.  
  	If this is not set only domain/enterprise admins will have access.  
  
 **-OrgUnitReset**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
  	Sets the OU that has reset access to the passwords.  
  	If this is not set only domain/enterprise admins will have access.  
  
 **-SecGroupRead**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
  	Sets the Security Group that has read access to the passwords stored in the computer AD object.  
  	If this is not set only domain/enterprise admins will have access.  
    
 **-SecGroupReset**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
  	Sets the Security Group that has reset access to the passwords.  
  	If this is not set only domain/enterprise admins will have access.  
    
 **-GPOName**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
  	Sets the GPO name when importing/creating the GPO.  
  	If this is not set the GPO name will be Deploy-LAPS.  
     
 **-DownloadURL**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
   Allows you to download LAPS from another source. This may be useful if a newer version comes out you can modify the script rather    
   than rely on me to update it.  
     
 **-DistributiveFileName**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
   Allows you to change the name of the executable that is downloaded. This could be useful where there is a SRP rule allowing a certain 
   filename in a certain directory.  
     
 **-InstallLogFileName**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
   Allows you to change the name of the installation log file name. This does not allow changing of the directory.  
     
 **-TranscriptFileName**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
   Allows you to change the name of the transcript. This does not allow changing of the directory.  
  
 **-SMBShareName**  
  	_Not Required._  
   _Pipeline not accepted._   
   _Named Position._   
   _No Wildcards._   
   Allows you to change the name of the SMB share created on the computer you run the script on to share out the executable for the 
   client computers to access.  


-------------------

Exmaple usage
-------------- 
    Requires admin shell.
  `.\Deploy-LAPS.ps1`
        This will run the script with all the defaults and prompting for required information as you go.
    
  `.\Deploy-LAPS.ps1 -GPOName "My Awesome LAPS GPO"`
  	    This will run the script with defaults but set the GPOname to "My Awesome LAPS GPO"
   
 
