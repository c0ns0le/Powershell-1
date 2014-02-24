CLEAR

$Server = "RCSDVDB1"

function storage
{
 param([string]$CompName)
 Get-WmiObject win32_volume -ComputerName $CompName -Filter "DriveType=3 AND Label <> 'System Reserved' AND DriveLetter IS NOT NULL" | Sort DriveLetter | ForEach-Object{"{0}, {1} - {2}gb" -f $_.Name,$_.Label,([Math]::Truncate($_.Capacity/1GB))}
}

function installedSW
{
 param([string]$CompName)
 Get-WmiObject win32_Product -ComputerName $CompName | Sort Name | ForEach-Object{"{0}" -f $_.Name}
}

function SQLinfo
{
	param([string]$CompName,[string]$OpRequest)
	
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
	
		while($sqlReaderAdmins.Read()) {$sqlReaderAdmins["name"]}
	
		$sqlReaderAdmins.Close()
	  }
  }
  
  $sqlConn.Close()
}

function writeList
{
 param([string[]]$inputFunction,
 		[string]$bookmarkName)
 
 # Build array of storage values; count total objects in array
 Write-Host "...building array"
 # $inputFunction
 $objectArrayCount = $inputFunction.Count
 
 # Locate the 'storage' Bookmark within the Word document
 Write-Host "...building range object(s)"
 $ObjRange = $wordDoc.Bookmarks.Item($bookmarkName).Range
 
 # Add a row for every object within $storageArray
 # This is designed to add one less row than there are drives on the server.
 # This is done so because the template starts with one row already available
 Write-Host "...adding rows"
 for ($i=1; $i -lt ($objectArrayCount-1); $i++)
 {
  $Rows = $ObjRange.Rows.Add()
  Write-Host ".........row " $i " added"
 }
 
 # Set table coordinates of storage bookmark
 Write-Host "...pinpointing coordinates"
 $CoordX = $ObjRange.Information(14) # Gets the ROW number for the storage Bookmark
 $CoordY = $ObjRange.Information(17) # Gets the COL number for the storage Bookmark
 
 Write-Host "...adding values"
 foreach ($object in $inputFunction)
 {
  # Add each drive to the document, starting at ($storageX,$storageY) and incrementing as needed
  $Table = $wordDoc.Tables.Item(1)
  $Table.Cell($CoordX,$CoordY).Range.Text = $object
  Write-Host "........." $object" added"
  $CoordX++
 }
 
}

# Create Microsoft Word ComObject
$msWord = New-Object -Com Word.Application

# Check path for validity; error out if the file does not exist
$docTemplate = "C:\Users\waclawskij\Desktop\Scripts\Documentation\Documents\BLANKdb.doc"
if (! (test-path $docTemplate))
{
  throw "$($docTemplate) is not a valid path"
}

# Open template and set as active document
Write-Host "== STARTING =="
$wordDoc = $msWord.Documents.Open($docTemplate)
$wordDoc.Activate()

# Write Storage Data
$func = storage -CompName $Server
writeList -inputFunction $func -bookmarkName "storage"

# Write Software Data
$func = installedSW -CompName $Server
writeList -inputFunction $func -bookmarkName "installedsw"

# Write SQL Admin Data
$func = SQLinfo -CompName $Server -OpRequest 'Admins'
writeList -inputFunction $func -bookmarkName "sql_admins"

# Save the file as testing.doc
Write-Host "...Saving"
$filename = "C:\Users\waclawskij\Desktop\Scripts\Documentation\Output\testing.doc"
$wordDoc.SaveAs([REF]$filename)
$wordDoc.Close()
Write-Host "/* Saved! *\"