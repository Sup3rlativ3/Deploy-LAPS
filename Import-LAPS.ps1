#Requires -Version 3 -RunAsAdministrator
#Requires -Modules AdmPwd.PS, ActiveDirectory, GroupPolicy

[CmdletBinding()]
param ([string[]]$ComputerOU,
	[string[]]$OrgUnitRead,
	[string[]]$OrgUnitReset,
	[string]$SecGroupRead,
	[string]$SecGroupReset,
	[string]$WorkFolderPath = (Join-Path -Path $env:homedrive -ChildPath 'Shadow\LAPS'),
	[string]$SMBShareName = 'LAPS$',
	[string]$TranscriptFileName = 'Script.log',
	[string]$InstallLogFileName = 'Install.log',
	[string]$GPOName = 'Deploy-LAPS',
	[string]$DownloadURL = 'https://download.microsoft.com/download/C/7/A/C7AAD914-A8A6-4904-88A1-29E657445D03/LAPS.x64.msi',
	[string]$DistributiveFileName = 'LAPSx64.msi'
)

# Find NetBIOS and FQDN names of the domain.
$NetBIOSName = (Get-ADDomain).NetBIOSName
$FQDN = (Get-ADDomain).DNSRoot


IF (!(Test-Path -Path $WorkFolderPath))
{
	New-Item -Path $WorkFolderPath -ItemType Directory -Force
}

# Start a transcipt of the script.
Start-Transcript -Path (Join-Path -Path $WorkFolderPath -ChildPath $TranscriptFileName)

# Create a share on the DC for the software. Add read access for domain computers to the share.
New-SmbShare -Name $SMBShareName -Path $WorkFolderPath -ReadAccess "$NetBIOSName\Domain Computers" -FullAccess "$NetBIOSName\Domain Admins"

IF (Test-Path -Path $WorkFolderPath)
{
	Copy-Item -Path .\ -Destination $WorkFolderPath -Recurse
}


# Download and install MS LAPS software.
$InstallationFilePath = (Join-Path -Path $WorkFolderPath -ChildPath $DistributiveFileName)
$DownloadStartTime = Get-Date

Invoke-WebRequest -Uri $DownloadURL -OutFile $InstallationFilePath
Write-Verbose -Message ('Time taken: {0} second(s)' -f ((Get-Date) - $DownloadStartTime).Seconds)

# Install management software on the DC and write a verbose log of the install.
Start-Process -FilePath 'msiexec' -ArgumentList ('/i {0} /passive /l*v "{1}" ADDDEFAULT=ALL' -f $InstallationFilePath, (Join-Path -Path $WorkFolderPath -ChildPath $InstallLogFileName))

# Import required PowerShell cmdlet.
# Update AD schema to accomodate the new fields to store password data.
Import-Module -Name AdmPwd.PS
Update-AdmPwdADSchema

# Check if a computer OU was provided by the parameter and act accordingly.
if (!$ComputerOU)
{
	# Get the default computer OU.
	# Reference https://support.microsoft.com/en-us/kb/324949
	$OUQuery = [adsisearcher]'(&(objectclass=domain))'
	$OUQuery.SearchScope = 'base'
	$OUQuery.FindOne().properties.wellknownobjects | ForEach-Object {
		if ($_ -match '^B:32:AA312825768811D1ADED00C04FD8D5CD:(.*)$')
		{
			$ComputerOU = $Matches[1]
			Write-Verbose -Message $ComputerOU
		}
	}
}
foreach ($OrgUnit in $ComputerOU)
{
	Set-AdmPwdComputerSelfPermission -Identity $OrgUnit
}

# Configure who can read the password. By default only domain/enterprise admins can.
if (!$OrgUnitRead)
{
	$OrgUnitRead = $ComputerOU
}
if ($SecGroupRead) {
	foreach ($OrgUnit in $OrgUnitRead)
	{
		Set-AdmPwdReadPasswordPermission -Identity $OrgUnit -AllowedPrincipals $SecGroupRead
	}
}

# Configure who can force a password change. By default only domain/enterprise admins can.
if (!$OrgUnitReset)
{
	$OrgUnitReset = $ComputerOU
}
if ($SecGroupReset) {
	foreach ($OrgUnit in $OrgUnitReset)
	{
		Set-AdmPwdResetPasswordPermission -Identity $OrgUnit -AllowedPrincipals $SecGroupReset
	}
}


# Importing the GPO
# Reference https://gallery.technet.microsoft.com/Migrate-Group-Policy-2b5067d8#content
Set-Location -Path $WorkFolderPath
Import-Module -Name GroupPolicy
Import-Module -Name ActiveDirectory

# Change variables in the GPO to suit environment using a GPO migration table.
$XML = [xml](Get-Content -Path (Join-Path -Path $WorkFolderPath -ChildPath 'MigrationTable.migtable'))
$Destination = $XML.MigrationTable.Mapping.ChildNodes

foreach ($node in $Destination)
{
	if ($node.'#text' -eq 'ENTERPRISEADMINS')
	{
		$Node.innertext = "Enterprise Admins@$FQDN"
	}
	if ($node.'#text' -eq 'DOMAINADMINS')
	{
		$Node.innertext = "Domain Admins@$FQDN"
	}
	if ($node.'#text' -like '\\SERVERNAME\*')
	{
		$Node.innertext = Join-Path -Path (Join-Path -Path (Join-Path -Path '\\' -ChildPath $env:computername) -ChildPath $SMBShareName) -ChildPath $DistributiveFileName
	}
}

$MigrationTableProcessedFilePath = Join-Path -Path $WorkFolderPath -ChildPath 'LAPS.migtable'
$XML.Save($MigrationTableProcessedFilePath)

#Import the actual GPO            
#http://serverfault.com/questions/491505/powershell-copy-gpo-failing-with-hresult-0x8007000d
Import-GPO -CreateIfNeeded -path $WorkFolderPath -BackupId '{4178CB42-3A58-445C-A46E-9CD8338C9FA5}' -TargetName $GPOName -MigrationTable $MigrationTableProcessedFilePath