Import-Module ActiveDirectory
Import-Module PSRemoteRegistry

function Get-Domain_OEC
{
 param([string]$CompName)
 $domain = Get-WmiObject -Query "SELECT Domain from Win32_ComputerSystem" -ComputerName $CompName #-Credential $cred
 $domain.Domain
}

function Get-OSVersion_OEC
{
 param([string]$CompName)
 $osver = Get-WmiObject -Query "SELECT Caption from Win32_OperatingSystem" -ComputerName $CompName #-Credential $cred
 $osver.Caption
}

function Get-CommisionDate_OEC
{
 param([string]$CompName)
 $CommDate = Get-ADComputer -Filter {Name -like $CompName} -Property *
 $CommDate.whenCreated
}

function Get-Processor_OEC
{
 param([string]$CompName)
 $processors = Get-WmiObject Win32_Processor -ComputerName $CompName
 
 foreach ($processor in $processors)
 {
 	($processor.Name) + "; " -replace "\s+"," "
 }
 
}

function Get-Memory_OEC
{
 param([string]$CompName)
 $memory = Get-WmiObject -Query "SELECT TotalPhysicalMemory from Win32_ComputerSystem" -ComputerName $CompName #-Credential $cred
 "{0:n2}" -f ([Math]::Truncate($memory.totalphysicalmemory / 1MB)) + " GB"
}

function Get-Storage_OEC
{
 param([string]$CompName)
 Get-WmiObject win32_volume -ComputerName $CompName -Filter "DriveType=3 AND Label <> 'System Reserved' AND DriveLetter IS NOT NULL" | Sort-Object DriveLetter | ForEach-Object{"{0}, {1} - {2}gb" -f $_.Name,$_.Label,([Math]::Truncate($_.Capacity/1GB))}
}

function Get-CanonicalName_OEC
{
 param([string]$CompName)
 $cname = Get-ADComputer -Filter {Name -like $CompName} -Property *
 $cname.CanonicalName
}

function Get-IPAddr_OEC
{
 param([string]$CompName)
 $wmi = Get-WmiObject -Query "SELECT * from Win32_NetworkAdapterConfiguration WHERE IPEnabled='true'" -ComputerName $CompName #-Credential $cred
 foreach ($nic in $wmi)
 {
  $nic.IPAddress
 }
}

function Get-Subnet_OEC
{
 param([string]$CompName)
 $wmi = Get-WmiObject -Query "SELECT * from Win32_NetworkAdapterConfiguration WHERE IPEnabled='true'" -ComputerName $CompName #-Credential $cred
 foreach ($nic in $wmi)
 {
  $nic.IPSubnet
 }
}

function Get-Gateway_OEC
{
 param([string]$CompName)
 $wmi = Get-WmiObject -Query "SELECT * from Win32_NetworkAdapterConfiguration WHERE IPEnabled='true'" -ComputerName $CompName #-Credential $cred
 foreach ($nic in $wmi)
 {
  $nic.DefaultIPGateway
 }
}

function Get-DNS_OEC
{
 param([string]$CompName)
 $wmi = Get-WmiObject -Query "SELECT * from Win32_NetworkAdapterConfiguration WHERE IPEnabled='true'" -ComputerName $CompName #-Credential $cred
 $wmi.DNSServerSearchOrder
}

function Get-CritServices_OEC
{
 param([string]$CompName,[string]$reqData)
 
 Switch ($reqData)
 {
	"names"{
			foreach ($service in $services)
			{
				Get-WmiObject win32_Service -ComputerName $computer | Sort Name | Where-Object {$_.Name -like $service} | ForEach-Object {"{0} - {1}" -f $_.DisplayName, $_.Name}
			}
		   }
	"runas"{
			foreach ($service in $services)
			{
				Get-WmiObject win32_Service -ComputerName $computer | Sort Name | Where-Object {$_.Name -like $service} | ForEach-Object {"{0}" -f $_.StartName}
			}
		   }
	"start"{
			foreach ($service in $services)
			{
				Get-WmiObject win32_Service -ComputerName $computer | Sort Name | Where-Object {$_.Name -like $service} | ForEach-Object {"{0}" -f $_.StartMode}
			}
		   }
	default {"'$reqData' is not a valid option.  Please choose either 'names', 'runas', or 'start'."}
 }
}

function Get-InstalledSW_OEC
{
 param([string]$CompName,[string]$reqData)
 # Get-WmiObject win32_Product -ComputerName $CompName | Sort Name | ForEach-Object{"{0}" -f $_.Name}
 
 Switch ($reqData)
 {
	"names"{
				Get-WmiObject win32_Product -ComputerName $CompName | Sort Name | ForEach-Object{"{0}" -f $_.Name}
		   }
	"version"{
				Get-WmiObject win32_Product -ComputerName $CompName | Sort Name | ForEach-Object{"{0}" -f $_.Version}
		   }
	default {"'$reqData' is not a valid option.  Please choose either 'names', 'runas', or 'start'."}
 }
}

function Get-SQLinfo_OEC
{
	param([string]$CompName,[string]$OpRequest)
	
	$Hive = [Microsoft.Win32.RegistryHive]“LocalMachine”;
	$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive,$CompName);
	$ref = $regKey.OpenSubKey("SOFTWARE\Microsoft\Microsoft SQL Server\MSSQLServer");
	
	if(!$ref)
	{
		"$CompName does not host any Microsoft SQL Server instances"
	}
	else
	{	
		#Create SQL Connection
		$sqlConn=New-Object System.Data.SqlClient.SqlConnection("Server=" + $CompName + ".oec.oeconnection.com;Database=master;Integrated Security=SSPI;")
		$sqlConn.Open()	
		
		switch($OpRequest)
		{
		  "Instance"
		  {
			# Build Query
			$sqlCommBasic=$sqlConn.CreateCommand()
			$sqlCommBasic.CommandText	= "DECLARE @GetInstName NVARCHAR(64),@InstName NVARCHAR(64),@Edition NVARCHAR(64)" +
										",@Version NVARCHAR(16) " +
										"SET @GetInstName = CONVERT(NVARCHAR,SERVERPROPERTY('InstanceName')) " +
										"SET @Edition = CONVERT(NVARCHAR,SERVERPROPERTY('Edition')) " +
										"SET @Version = CONVERT(NVARCHAR,SERVERPROPERTY('ProductVersion')) " +
										"IF @GetInstName IS NULL BEGIN SET @InstName = '<default>' END " +
										"ELSE SET @InstName = @GetInstName " + 
										"SELECT @InstName AS instance,@Edition as edition,@Version as version"	
		
			# Execute Query
			$sqlReaderBasic 			= $sqlCommBasic.ExecuteReader()
		
			# Fill Variables
			while($sqlReaderBasic.Read())
			{
				$InstanceName			= $sqlReaderBasic["instance"]
			}
		
			# Close Connection
			$sqlReaderBasic.Close()
			
			return $InstanceName
		  }
		  "Edition"
		  {
			# Build Query
			$sqlCommBasic=$sqlConn.CreateCommand()
			$sqlCommBasic.CommandText	= "DECLARE @GetInstName NVARCHAR(64),@InstName NVARCHAR(64),@Edition NVARCHAR(64)" +
										",@Version NVARCHAR(16) " +
										"SET @GetInstName = CONVERT(NVARCHAR,SERVERPROPERTY('InstanceName')) " +
										"SET @Edition = CONVERT(NVARCHAR,SERVERPROPERTY('Edition')) " +
										"SET @Version = CONVERT(NVARCHAR,SERVERPROPERTY('ProductVersion')) " +
										"IF @GetInstName IS NULL BEGIN SET @InstName = '<default>' END " +
										"ELSE SET @InstName = @GetInstName " + 
										"SELECT @InstName AS instance,@Edition as edition,@Version as version"	
		
			# Execute Query
			$sqlReaderBasic 			= $sqlCommBasic.ExecuteReader()
		
			# Fill Variables
			while($sqlReaderBasic.Read())
			{
				$Edition				= $sqlReaderBasic["edition"]
				$Version				= $sqlReaderBasic["version"]
			}
		
			# Close Connection
			$sqlReaderBasic.Close()
			
			$Output = $Edition + ", " + $Version
			
			return $Output
		  }
		  "Login"
		  {
			# Get Server Login Information
			$sqlCommLogin=$sqlConn.CreateCommand()
			$sqlCommLogin.CommandText	= "SELECT CASE SERVERPROPERTY('IsIntegratedSecurityOnly') " + 
										"WHEN 1 THEN 'Windows Authentication' " + 
										"WHEN 0 THEN 'Windows and SQL Server Authentication' " + 
										"END as [Mode]"
		
			# Execute Query
			$sqlReaderLogin 			= $sqlCommLogin.ExecuteReader()
		
			# Fill Variables
			while($sqlReaderLogin.Read())
			{
				$Login					= $sqlReaderLogin["Mode"]
			}
		
			# Close Connection
			$sqlReaderLogin.Close()
			
			return $Login
		  }
		  "Admins"
		  {
		  # Get SQL Server admin users
			$sqlCommAdmins=$sqlConn.CreateCommand()
		
			$sqlCommAdmins.CommandText	="SELECT name,IS_SRVROLEMEMBER('sysadmin', name) AS [Admin] " +
										"FROM sys.server_principals " +
										"WHERE IS_SRVROLEMEMBER('sysadmin', name) = 1 " +
										"AND name NOT LIKE 'NT%' AND name NOT LIKE '%SQL%' ORDER BY name"
		
			$sqlReaderAdmins 			= $sqlCommAdmins.ExecuteReader()
		
			while($sqlReaderAdmins.Read()) 
			{
				"`n" + $sqlReaderAdmins["name"]
			}
		
			$sqlReaderAdmins.Close()
		  }
		  "Databases"
		  {
		  # Get SQL Server User Databases
			$sqlCommDB=$sqlConn.CreateCommand()

			$sqlCommDB.CommandText	="SELECT name " +
									"FROM master.sys.databases " +
									"WHERE database_id > 4"

			$sqlReaderDB 			= $sqlCommDB.ExecuteReader()

			while($sqlReaderDB.Read()) 
			{
				"`n" + $sqlReaderDB["name"]
			}

			$sqlReaderDB.Close()
		  }
	  }
	  
	  $sqlConn.Close()
	}
}

function Write-List_OEC
{
 param([string[]]$inputFunction,
 		[string]$bookmarkName)
 
 # Build array of values; count total objects in array
 Write-Host "......building array"
 # $inputFunction
 $objectArrayCount = $inputFunction.Count
 
 # Locate the Bookmark within the Word document
 Write-Host "......building range object(s)"
 $ObjRange = $wordDoc.Bookmarks.Item($bookmarkName).Range
 
 # Add a row for every object within $storageArray
 # This is designed to add one less row than there are items in the collection.
 # This is done so because the template starts with one row already available
 Write-Host "......adding rows"
 
 if ($bookmarkName -like "*runas" -or $bookmarkName -like "*start" -or $bookmarkName -like "*version")
 	{
		Write-Host "............This item type does not need new rows."
	}
 else
 	{
		for ($i=1; $i -lt ($objectArrayCount-1); $i++)
		 {
		  $Rows = $ObjRange.Rows.Add()
		  Write-Host "............row " $i " added"
		 }
	}
 # Set table coordinates of storage bookmark
 Write-Host "......pinpointing coordinates"
 $CoordX = $ObjRange.Information(14) # Gets the ROW number for the Bookmark
 $CoordY = $ObjRange.Information(17) # Gets the COL number for the Bookmark
 
 Write-Host "......adding values"
 foreach ($object in $inputFunction)
 {
  # Add each item to the document, starting at ($CoordX,$CoordY) and incrementing as needed
  $Table = $wordDoc.Tables.Item(1)
  $Table.Cell($CoordX,$CoordY).Range.Text = $object
  Write-Host "............" $object" added"
  $CoordX++
 }
 
}

function Get-FolderSize_OEC
{
	param([string]$dir)
	$rawSize = Get-ChildItem $dir -recurse | Measure-Object -property length -sum	
	
	"{0:N2}" -f ($rawSize.sum / 1GB) + " GB"
}

function Get-PCUpTime_OEC ($computer)
{
	$lastboottime = (Get-WMIObject -Class Win32_OperatingSystem -computername $computer).LastBootUpTime
	$sysuptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($lastboottime)
	Write-Host "$computer has been up for: " $sysuptime.days "Days" $sysuptime.hours "Hours" $sysuptime.minutes "Minutes" $sysuptime.seconds "Seconds"
} 

function Get-Process_OEC ($computer,$pname)
{
	Get-WmiObject win32_process -ComputerName $computer | Where-Object {$_.name -like "*$pname*"} | Select Name,ProcessId
}

function End-RemoteProcess_OEC ($computer,$pid)
{
	(Get-WmiObject win32_process -ComputerName $computer -Filter "ProcessId = $pid").Terminate()
}

function Get-RegInfo_OEC
{
	param([string]$Server,[string]$keyPath,[string]$value)
	
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
}