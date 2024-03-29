CLEAR

Import-Module PSRemoteRegistry
Import-Module ActiveDirectory
Import-Module DocumentationTools

# Set file with server names
$hostInputFile = "D:\Development\Powershell\Documentation\Documents\servers.txt"

# Check existence
if (! (test-path $hostInputFile))
{
 throw "$($hostInputFile) is not a valid path."
}

# Build server list array
$computers = Get-Content $hostInputFile

foreach ($computer in $computers)
{
	$Hive = [Microsoft.Win32.RegistryHive]“LocalMachine”;
	$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive,$computer);
	$ref = $regKey.OpenSubKey("SOFTWARE\Microsoft\Microsoft SQL Server\MSSQLServer");
	
	if(!$ref)
	{
		##"$CompName does not host any Microsoft SQL Server instances"
	}
	else
	{	
		"`n==== $computer ===="
		## Instance(s)
		$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Microsoft\Microsoft SQL Server\" -value "InstalledInstances" 
		Write-Host ">> Instance(s) :: $cmd"
		
		## Version
		$cmd=Get-SQLinfo_OEC -CompName $computer -OpRequest Edition
		Write-Host ">> Edition :: $cmd"
		
		## StartupParameters
		$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Microsoft\Microsoft SQL Server\" -value "SQLArg*" 
		Write-Host ">> Startup Parameters :: $cmd"
		
		## SQL Directories
		$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Microsoft\Microsoft SQL Server\" -value "SqlProgramDir"
		Write-Host ">> Shared Features Directory :: $cmd"
		$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Microsoft\Microsoft SQL Server\" -value "SQLPath"
		Write-Host ">> Root Directory :: $cmd"
		$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Microsoft\Microsoft SQL Server\" -value "DefaultData"
		Write-Host ">> Default Data Directory :: $cmd"
		$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Microsoft\Microsoft SQL Server\" -value "DefaultLog"
		Write-Host ">> Default Log Directory :: $cmd"
		#$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Microsoft\Microsoft SQL Server\" -value "<value_here>"
		#Write-Host ">> TempDB Directory :: $cmd"
		$cmd=Get-RegInfo_OEC -Server $computer -keyPath "SOFTWARE\Red Gate\SQL Backup\BackupSettingsGlobal" -value "BackupFolder"
		Write-Host ">> Backup Directory :: $cmd"
		
		## User Databases
		$cmd=Get-SQLinfo_OEC -CompName $computer -OpRequest Databases
		Write-Host ">> Databases :: $cmd"
		
		## Authentication Type
		$cmd=Get-SQLinfo_OEC -CompName $computer -OpRequest Login
		Write-Host ">> Authentication :: $cmd"
		
		## Admins
		$cmd=Get-SQLinfo_OEC -CompName $computer -OpRequest Admins
		Write-Host ">> Administrators :: $cmd"
	}
}