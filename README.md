Deploy-LAPS
=================

This script and accompanying files will allow system administrators to automatically deploy Microsoft Local Administrator Password Solution (LAPS). Please visit https://technet.microsoft.com/en-us/mt227395.aspx for more information about LAPS.



 Parameters
 -------------- 
 **-ComputerOU**  
  	Not Required.  
  	This is the OU that LAPS will assign permissions to update it's own AD object to write the password when changed.
  	If this is not set it will pull the default computer OU from AD.  
  
 **-OrgUnitRead**  
  	Not Required  
  	Sets the OU that has read access to the passwords stored in the computer AD object.  
  	If this is not set only domain/enterprise admins will have access.  
  
 **-OrgUnitReset**  
  	Not Required  
  	Sets the OU that has reset access to the passwords.  
  	If this is not set only domain/enterprise admins will have access.  
  
 **-SecGroupRead**  
  	Not Required  
  	Sets the Security Group that has read access to the passwords stored in the computer AD object.  
  	If this is not set only domain/enterprise admins will have access.  
    
 **-SecGroupReset**  
  	Not Required  
  	Sets the Security Group that has reset access to the passwords.  
  	If this is not set only domain/enterprise admins will have access.  
    
 **-GPOName**  
  	Not Required  
  	Sets the GPO name when importing/creating the GPO.  
  	If this is not set the GPO name will be Deploy-LAPS.  
     
 **-DownloadURL**  
   Not Required  
   Allows you to download LAPS from another source. This may be useful if a newer version comes out you can modify the script rather    
   than rely on me to update it.  
     
 **-DistributiveFileName**  
   Not Required  
   Allows you to change the name of the executable that is downloaded. This could be useful where there is a SRP rule allowing a certain 
   filename in a certain directory.  
     
 **-InstallLogFileName**  
   Not Required  
   Allows you to change the name of the installation log file name. This does not allow changing of the directory.  
     
 **-TranscriptFileName**  
   Not Required  
   Allows you to change the name of the transcript. This does not allow changing of the directory.  
  
 **-SMBShareName**  
   Not Required  
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
   
 
