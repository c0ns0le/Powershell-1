CLEAR

## Set IIS Log File Directory Variable
$logFileDir_Top = (Get-ItemProperty "IIS:\Sites\Microsoft Dynamics CRM" -name logFile.directory).Value

WRITE-HOST $logFileDir_Top