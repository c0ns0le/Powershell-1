CLEAR

$srvr = “RCSDVDB1”

$Hive = [Microsoft.Win32.RegistryHive]“LocalMachine”;
$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive,$srvr);

$ref = $regKey.OpenSubKey(“SOFTWARE\Red Gate”);
	if (!$ref)
	{
		"Does not exist"
	}
	else
	{
		"Exists"
	}