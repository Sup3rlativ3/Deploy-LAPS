#requires -version 3 -RunAsAdministrator

param ( [string]$ComputerOU, 
        [string]$OrgUnitRead, 
        [string]$OrgUnitReset, 
        [string]$SecGroupRead, 
        [string]$SecGroupReset,
        [string]$GPOName
      )

# Find Netbios name of domain.
$NetBIOSName = (Get-ADDomain).NetBIOSName
$FQDN = (Get-ADDomain).DNSRoot


IF(!(Test-Path $env:homedrive\Shadow\LAPS))
    {
        New-Item "$env:homedrive\Shadow\LAPS" -ItemType Directory -Force
    }

# Start a transcipt of the script.
Start-Transcript "C:\Shadow\LAPS\Script.log"

# Create share on the DC for the software. Add domain computers read access to share.
New-SmbShare -Name "LAPS$" -Path "$env:homedrive\Shadow\LAPS\" -ReadAccess "$NetBIOSName\Domain computers"
      
IF(!(Test-Path $env:homedrive\Shadow\LAPS))
    {
        Copy-Item -Path .\ -Destination "$env:homedrive\Shadow\LAPS" -Recurse
    }

IF (!($GPOName))
    {
        $GPOName = "Deploy-LAPS"
    }


# Download and install the MS LAPS software.
$url = "https://download.microsoft.com/download/C/7/A/C7AAD914-A8A6-4904-88A1-29E657445D03/LAPS.x64.msi"
$output = "$env:homedrive\Shadow\LAPS\LAPSx64.msi"
$start_time = Get-Date

Invoke-WebRequest -Uri $url -OutFile $output
Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"

# Install mgmt software on DC and write a verbose log of the install.
msiexec /i "$env:homedrive\Shadow\LAPS\LAPSx64.msi" /passive /l*v "$env:homedrive\Shadow\LAPS\Install.log" ADDDEFAULT=ALL

# Import the new powershell cmdlets.
# Update the AD schema to accomodate the new field for the password.

Import-Module AdmPwd.PS
Update-AdmPwdADSchema

# Get the default computer OU.
# Reference https://support.microsoft.com/en-us/kb/324949
$OUQuery = [adsisearcher]'(&(objectclass=domain))'
$OUQuery.SearchScope = 'base'
$OUQuery.FindOne().properties.wellknownobjects | ForEach-Object {
    if ($_ -match '^B:32:AA312825768811D1ADED00C04FD8D5CD:(.*)$')
            {
                $Matches[1]
            }
        }


# Check if a computer OU was provided by the parameter and act accordingly.
if (!$ComputerOU)
    {
        Set-AdmPwdComputerSelfPermission -OrgUnit $Matches[1]
    }
    else
        {
            Set-AdmPwdComputerSelfPermission -OrgUnit $ComputerOU
        }

# Configure who can read the attribute. By default only domain/enterprise admins can.
if ($OrgUnitRead)
    {
        Set-AdmPwdReadPasswordPermission -OrgUnit $OrgUnitRead
    }
if ($SecGroupRead)
    {
        Set-AdmPwdReadPasswordPermission -AllowedPrincipals $SecGroupRead
    }

# Configure who can force a password change. By default only domain/enterprise admins can.
if ($OrgUnitReset)
    {
        Set-AdmPwdResetPasswordPermission -OrgUnit $OrgUnitReset
    }
if ($SecGroupReset)
    {
        Set-AdmPwdResetPasswordPermission -AllowedPrincipals $SecGroupReset
    }


# Importing the GPO
# Reference https://gallery.technet.microsoft.com/Migrate-Group-Policy-2b5067d8#content
Set-Location "$env:homedrive\Shadow\LAPS"
Import-Module GroupPolicy            
Import-Module ActiveDirectory            
# change vars in csv to suit environment
$MigrationMap = Import-Csv .\MigrationMap.csv 

$Output = ForEach ($r in $MigrationMap) 
    {
        if ($r.Destination -like "DestDomainFQDN")
            {
                $r.Destination = "$FQDN"
                $r
            }
            elseif ($r.Destination -like "DestDomainNetBIOS")
                {
                    $r.Destination = "$NetBIOSName"
                    $r
                }
            elseif ($r.Destination -like "DestDomainFQDNUNC")
                {
                    $r.Destination = "\\$FQDN\"
                    $r
                }
            elseif ($r.Destination -like "DestDomainNetBIOSUNC")
                {
                    $r.Destination = "\\$NetBIOSName\"
                    $r
                }
    }
    
$Output | Export-CSV .\MigrationMap.csv -NoTypeInformation

            
        
 
#Import the actual GPO            
Import-GPO -CreateIfNeeded -path $env:homedrive\Shadow\LAPS\ -BackupId "{4178CB42-3A58-445C-A46E-9CD8338C9FA5}" -TargetName $GPOName -MigrationTable $env:homedrive\Shadow\LAPS\MigrationMap.csv
    