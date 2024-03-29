CLEAR

Import-Module ActiveDirectory
Import-Module DocumentationTools

# specify full path to the list of servers and services you want to scan.
$hostInputFile = "D:\Development\Powershell\Documentation\Documents\servers_testing.txt"

if (! (test-path $hostInputFile))
{
 throw "$($hostInputFile) is not a valid path."
}

$computers = Get-Content $hostInputFile

foreach($computer in $computers)
{
	####################################################################
	####  This section is used to set all variables required for final
	####  insert at the end of the script.
	####################################################################
	
	# Build variables requiring NO queries
	$connString="Server=" + $computer + ".oec.oeconnection.com;Database=master;Integrated Security=SSPI;"
	$FQDN=$computer + ".oec.oeconnection.com"
	
	# Query for Windows Edition
	$osVer = Get-OSVersion_OEC -CompName $computer
	
	# Query for Commision Date
	$commDate = Get-CommisionDate_OEC -CompName $computer
	
	# Query for Processor
	$processors = Get-Processor_OEC -CompName $computer
	
	# Query for Memory
	$memory = Get-Memory_OEC -CompName $computer
	
	# Create SQL Connection
	$sqlConn=New-Object System.Data.SqlClient.SqlConnection("Server=RCSCOMBIS.oec.oeconnection.com;Database=master;Integrated Security=SSPI;")
	$sqlConn.Open()
	
	# Create & Execute SQL Query
	$sqlCommBasic=$sqlConn.CreateCommand()
	$sqlCommBasic.CommandText	= "INSERT INTO [Testing].[dbo].[BISServers] " + 
								  "([serverName],[FQDN],[connString],[winEdition],[commision]) " + 
								  "VALUES " + 
								  "('$computer','$FQDN','$connString','$osVer','$commDate') "
	
	$sqlCommBasic.ExecuteNonQuery()
}