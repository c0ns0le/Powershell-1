CLEAR

Import-Module ActiveDirectory

function Get-ADObjectType
{
	param ($ObjName)
	
	$ObjInfo = Get-ADObject -LDAPFilter "(SamAccountName=$ObjName)"
	
	return $ObjInfo.ObjectClass
}

function Get-LocalAdmins 
{
	param ($CompName)
	
	$admins = Gwmi win32_groupuser –computer $CompName   
	$admins = $admins |? {$_.groupcomponent –like '*"Administrators"'}  
  	
	$admins |% {$_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul  
				$matches[1].trim('"') + “\” + $matches[2].trim('"')}
}

function SQLinfo
{
	param([string]$CompName)
	
	if ($CompName -eq "RCBISUTIL")
	{
		$sqlConn=New-Object System.Data.SqlClient.SqlConnection("Server=" + $CompName + ";Database=master;Integrated Security=SSPI;")
		$sqlConn.Open()
		
		$sqlCommAdmins=$sqlConn.CreateCommand()
		$sqlCommAdmins.CommandText	="SELECT name " + 
									"FROM master.dbo.syslogins " + 
									"WHERE sysadmin = 1 " + 
									"OR securityadmin = 1"
	
		$sqlReaderAdmins 			= $sqlCommAdmins.ExecuteReader()
	
		while($sqlReaderAdmins.Read()) {$sqlReaderAdmins["name"]}
	
		$sqlReaderAdmins.Close()
		return $Admins
		$sqlConn.Close()
	}
	else
	{
		$sqlConn=New-Object System.Data.SqlClient.SqlConnection("Server=" + $CompName + ";Database=master;Integrated Security=SSPI;")
		$sqlConn.Open()
		
		$sqlCommAdmins=$sqlConn.CreateCommand()
		$sqlCommAdmins.CommandText	="SELECT name,IS_SRVROLEMEMBER('sysadmin', name) AS [Admin] " +
									"FROM sys.server_principals " +
									"WHERE IS_SRVROLEMEMBER('sysadmin', name) = 1 " +
									"AND name NOT LIKE 'NT%' AND name NOT LIKE '%SQL%' ORDER BY name"
	
		$sqlReaderAdmins 			= $sqlCommAdmins.ExecuteReader()
	
		while($sqlReaderAdmins.Read()) {$sqlReaderAdmins["name"]}
	
		$sqlReaderAdmins.Close()
		return $Admins
		$sqlConn.Close()
	}
}


$hostInputFile = "C:\Users\waclawskij\Desktop\Admins\servers.txt"
if (! (test-path $hostInputFile))
{
 	throw "$($hostInputFile) is not a valid path."
}

$computers = Get-Content $hostInputFile

foreach ($computer in $computers)
{
	## Get and show users in local administrators group
	Write-Host "Probing "$computer"..."
	"== **" + $computer + "**, Local Administrators Group ==" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
	
	$AdminObjs = Get-LocalAdmins -CompName $computer
	foreach ($obj in $AdminObjs)
	{
		Write-Host "...Checking object class"
		if($obj.StartsWith("OEC\"))
		{
			$objRedux = $obj.Substring(4)
			if ((Get-ADObjectType -ObjName $objRedux) -eq 'group')
			{
				$obj + "*" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
			}
			else
			{
				$obj | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
			}
		}
		else
		{
			$obj | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
		}
	}
	
	if ($computer -like "*DB*" -or $computer -like "*UTIL" -or $computer -like "*SCOM*")
 	{
		Write-Host "Probing "$computer" SQL Settings..."
		# Get and show users with SysAdmin Role in SQL
		if($computer -eq "RCBISUTIL")
		{
			$SQLInstances = @("\MSSQLSERVER2008","")
			foreach ($instance in $SQLInstances)
			{
				$SQLServer = $computer + $instance
				switch ($SQLServer)
				{
					"RCBISUTIL"
					{
						Write-Host "..."$SQLServer"\MSSQLSERVER2000"
						"`n== " + $SQLServer + "\MSSQLSERVER2000, SQL Administrators ==" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
						$SQLAdmins = SQLinfo -CompName $SQLServer | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
						
						foreach ($sqlObj in $SQLAdmins)
						{
							if($sqlObj -like "OEC\*")
							{
								$sqlObjRedux = $sqlObj.Substring(4)
								if ((Get-ADObjectType -ObjName $sqlObjRedux) -eq 'group')
								{
									$sqlObj + "*" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
								}
								else
								{
									$sqlObj | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
								}
							}
							else
							{
								$sqlObj | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
							}
						}
					}
					"RCBISUTIL\MSSQLSERVER2008"
					{
						Write-Host "..."$SQLServer
						"`n== " + $SQLServer + ", SQL Administrators ==" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
						$SQLAdmins = SQLinfo -CompName $SQLServer | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
						
						foreach ($sqlObj in $SQLAdmins)
						{
							if($sqlObj -like "OEC\*")
							{
								$sqlObjRedux = $sqlObj.Substring(4)
								if ((Get-ADObjectType -ObjName $sqlObjRedux) -eq 'group')
								{
									$sqlObj + "*" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
								}
								else
								{
									$sqlObj | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
								}
							}
							else
							{
								$sqlObj | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
							}
						}
					}
				}	
			}
		}
		else
		{
			"`n== " + $computer + ", SQL Administrators ==" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
			SQLinfo -CompName $computer | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
		}
	}
	
	"`n" | Out-File "C:\Users\waclawskij\Desktop\Admins\admin_data.txt" -Append
}

Write-Host "COMPLETE"