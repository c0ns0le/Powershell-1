CLEAR

Import-Module PSRemoteRegistry
Import-Module ActiveDirectory
Import-Module DocumentationTools

$Server="RCSDVDB1"
$keyPath="SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQLServer\"
$value="SQLArg*"

$Hive = [Microsoft.Win32.RegistryHive]“LocalMachine”;
$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive,$Server);
$ref = $regKey.OpenSubKey($keyPath);

if (!$ref)
{
"Key does not exist"
}
else
{
	$RegEntry = $Server | Get-RegValue -Key $keyPath -Value $value -Recurse
	if (!$RegEntry)
	{
		"Value does not exist for server"
	}
	else
	{
		foreach ($val in $RegEntry)
		{
			$val.Data
		}
	}
}