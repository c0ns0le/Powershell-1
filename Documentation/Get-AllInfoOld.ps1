CLEAR
Import-Module ActiveDirectory

$Server = Read-Host "What server?"
#$Server			= "RCSDVDB1"
$Type			= 'computer'
################################
###### System Information ######
################################
$ComputerInfo	= Get-WmiObject win32_computersystem -ComputerName $Server
$ProcInfo		= Get-WmiObject win32_processor -ComputerName $Server
$ADInfo			= Get-ADComputer -Filter {Name -like $Server} -Property *
$NetInfo		= Get-WmiObject win32_NetworkAdapterConfiguration -ComputerName $Server -Filter IPEnabled=TRUE
#$Svcs			= Get-Service -ComputerName $Server #| Select DisplayName | Sort-Object DisplayName | Where-Object {$_.DisplayName -like "SQL Server Integration*" -or $_.DisplayName -like "SQL Server Reporting*" -or $_.DisplayName -like "SQL Server Analysis*"} | ft -HideTableHeaders

$Domain			= $ComputerInfo.Domain
$OSEdition		= $ADInfo.OperatingSystem
$OSServicePack	= $ADInfo.OperatingSystemServicePack
$OSVersion		= $ADInfo.OperatingSystemVersion
$CreateDate		= $ADInfo.whenCreated

$Proc			= $ProcInfo | Select-Object -ExpandProperty Name
$Memory			= $ComputerInfo | ForEach-Object {[Math]::Truncate($_.totalphysicalmemory / 1MB)}
$Storage		= Get-WmiObject win32_volume -ComputerName RCSDVDB1 -Filter "DriveType=3 AND Label <> 'System Reserved' AND DriveLetter IS NOT NULL" | Sort-Object DriveLetter | ForEach-Object{"{0}, {1} - {2}gb `n" -f $_.Name,$_.Label,([Math]::Truncate($_.Capacity/1GB))}

$CName			= $ADInfo.CanonicalName
$IPAddr			= $NetInfo.IPAddress
$Subnet			= $NetInfo.IPSubnet
$Gateway		= $NetInfo.DefaultIPGateway
$DNS			= $NetInfo.DNSServerSearchOrder

"========================="
"===Systems Information==="
"========================="
"=== Basic Information ==="
"Server: " 			+ $Server
"FQDN: " 			+ $Server + '.' + $Domain
"Windows Edition: " + $OSEdition + " " + $OSServicePack + ", Version: " + $OSVersion
"Commission Date: " + $CreateDate

"`n=== Hardware Details ==="
"Processor: " 		+ $Proc -replace '\s+', ' '
"Memory: " 			+ $Memory + " Mb"
"Storage: " 		+ $Storage

"=== Network Details ==="
"Domain: " 			+ $Domain
"Canonical Name: " 	+ $CName
"IP Address(es): " 	+ $IPAddr
"Subnet Mask: " 	+ $Subnet
"Default Gateway: " + $Gateway
"DNS Server(s): " 	+ $DNS

################################
##### Software Information #####
################################

$InstalledSW=Get-WmiObject win32_Product -ComputerName $Server | Select Name,Version | Sort Name | FT -HideTableHeaders

"========================="
"==Software  Information=="
"========================="
"`n=== Critical Services ==="

"`n=== Installed Software ==="
$InstalledSW

If ($Server -like "*DB*")
{	
	
	################################
	####### SQL  Information #######
	################################
	
	$sqlConn=New-Object System.Data.SqlClient.SqlConnection("Server=" + $Server + ".oec.oeconnection.com;Database=master;Integrated Security=SSPI;")
	$sqlConn.Open()
	
	# Get Basic SQL Server Information
	$sqlCommBasic=$sqlConn.CreateCommand()
	$sqlCommBasic.CommandText	= "DECLARE @GetInstName NVARCHAR(64),@InstName NVARCHAR(64),@Edition NVARCHAR(64)" +
									",@Version NVARCHAR(16) " +
									"SET @GetInstName = CONVERT(NVARCHAR,SERVERPROPERTY('InstanceName')) " +
									"SET @Edition = CONVERT(NVARCHAR,SERVERPROPERTY('Edition')) " +
									"SET @Version = CONVERT(NVARCHAR,SERVERPROPERTY('ProductVersion')) " +
									"IF @GetInstName IS NULL BEGIN SET @InstName = '<default>' END " +
									"ELSE SET @InstName = @GetInstName " + 
									"SELECT @InstName AS instance,@Edition as edition,@Version as version"

	$sqlReaderBasic 			= $sqlCommBasic.ExecuteReader()
	
	while($sqlReaderBasic.Read())
	{
		$InstanceName			= $sqlReaderBasic["instance"]
		$Edition				= $sqlReaderBasic["edition"]
		$Version				= $sqlReaderBasic["version"]
	}
	
	$sqlReaderBasic.Close()
	
	# Get Server Login Information
	$sqlCommLogin=$sqlConn.CreateCommand()
	$sqlCommLogin.CommandText	= "SELECT CASE SERVERPROPERTY('IsIntegratedSecurityOnly') " + 
										"WHEN 1 THEN 'Windows Authentication' " + 
										"WHEN 0 THEN 'Windows and SQL Server Authentication' " + 
										"END as [Mode]"

	$sqlReaderLogin 			= $sqlCommLogin.ExecuteReader()
	
	while($sqlReaderLogin.Read())
	{
		$Login					= $sqlReaderLogin["Mode"]
	}
	
	$sqlReaderLogin.Close()
	
	"========================="
	"=====SQL Information====="
	"========================="
	"`n=== Basic Information ==="	
	"Instance(s): " 			+ $InstanceName
	"Version: " 				+ $Edition + ", " + $Version
	
	"`n====== ServerSecurity ======"
	"Login Mode: " + $Login
	
	# Get SQL Server admin users
	$sqlCommAdmins=$sqlConn.CreateCommand()
	$sqlCommAdmins.CommandText	="SELECT name,IS_SRVROLEMEMBER('sysadmin', name) AS [Admin] " +
									"FROM sys.server_principals " +
									"WHERE IS_SRVROLEMEMBER('sysadmin', name) = 1 " +
									"AND name NOT LIKE 'NT%' AND name NOT LIKE '%SQL%' ORDER BY name"
	
	$sqlReaderAdmins 			= $sqlCommAdmins.ExecuteReader()
	"`n=== Administrative Users ==="
	while($sqlReaderAdmins.Read()) {$sqlReaderAdmins["name"]}
	$sqlReaderAdmins.Close()

	# Get all User Databases
	$sqlCommDB=$sqlConn.CreateCommand()
	$sqlCommDB.CommandText="SELECT name FROM sys.databases WHERE database_id > 6 ORDER BY name"

	$sqlReaderDB = $sqlCommDB.ExecuteReader()
	"`n=== User Databases ==="
	while($sqlReaderDB.Read()) {$sqlReaderDB["name"]}
	$sqlReaderDB.Close()
	
	$sqlConn.Close()
}
ElseIf($Server -like "*WEB*")
{
	"web here"
}