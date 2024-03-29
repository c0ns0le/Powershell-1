CLEAR
## Function creation
    function Get-FileSize_OEC
    {
    	param([string]$file)
    	$rawSize = Get-ChildItem $file -recurse | Measure-Object -property length -sum	
    	
    	"{0:N2}" -f ($rawSize.sum / 1GB) + " GB"
    }
## End function creation

## Create file variables
    $PROD_Data = '\\rcproddb4\E$\SQLData\NAVREPORTING\ApplicationIntegration.mdf'
    $PROD_Log = '\\rcproddb4\F$\SQLLogs\NAVREPORTING\ApplicationIntegration_log.ldf'

    $DEV_Data = '\\Rcsdvdb4\e$\SQLData\NAVREPORTING\ApplicationIntegration.mdf'
    $DEV_Log = '\\Rcsdvdb4\F$\SQLLogs\NAVREPORTING\ApplicationIntegration_log.ldf'
## End create file variables

## Get file sizes
    $PROD_Data_Size = Get-FileSize_OEC -file $PROD_Data
    $PROD_Log_Size = Get-FileSize_OEC -file $PROD_Log
    
    $DEV_Data_Size = Get-FileSize_OEC -file $DEV_Data
    $DEV_Log_Size = Get-FileSize_OEC -file $DEV_Log
## End get file sizes

## Write file sizes
    $msg = "`n`nFile Sizes on " + (Get-Date) + 
            "`n========================================================================" + 
            "`n>> Production, ApplicationIntegration (DATA) - " + $PROD_Data_Size + 
            "`n>> Production, ApplicationIntegration (LOG)  - " + $PROD_Log_Size + 
            "`n>> Development, ApplicationIntegration (DATA)- " + $DEV_Data_Size +
            "`n>> Development, ApplicationIntegration (LOG) - " + $DEV_Log_Size
            
    $msg >> "C:\Users\waclawskij\Desktop\output.txt"
## End write file sizes