CLEAR

$sqlConn=New-Object System.Data.SqlClient.SqlConnection("Server=RCSDVDB1.oec.oeconnection.com;Database=master;Integrated Security=SSPI;")
$sqlConn.Open()

$sqlCommAdmin=$sqlConn.CreateCommand()
$sqlCommAdmin.CommandText="SELECT name,IS_SRVROLEMEMBER('sysadmin', name) AS [Admin] FROM sys.server_principals WHERE IS_SRVROLEMEMBER('sysadmin', name) = 1 AND name NOT LIKE 'NT%' AND name NOT LIKE '%SQL%' ORDER BY name"

$sqlReaderAdmin = $sqlCommAdmin.ExecuteReader()

"=== Administrative Users ==="
while($sqlReaderAdmin.Read()) {$sqlReaderAdmin["name"] + " -- " + $sqlReaderAdmin["Admin"]}

$sqlReaderAdmin.Close()

$sqlCommDB=$sqlConn.CreateCommand()
$sqlCommDB.CommandText="SELECT name FROM sys.databases WHERE database_id > 6 ORDER BY name"

$sqlReaderDB = $sqlCommDB.ExecuteReader()

"`n=== User Databases ==="
while($sqlReaderDB.Read()) {$sqlReaderDB["name"]}

$sqlConn.Close()