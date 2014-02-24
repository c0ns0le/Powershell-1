CLEAR

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
  }
  
  $sqlConn.Close()
}

SQLinfo -CompName 'RCSDVDB1' -OpRequest 'Instance'
SQLinfo -CompName 'RCSDVDB1' -OpRequest 'Edition'
SQLinfo -CompName 'RCSDVDB1' -OpRequest 'Login'