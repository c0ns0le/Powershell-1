PARAM
(
    [string]$ServerName
)

if($ServerName -like "*WEB*")
{
    ## Load existing XML configuration file
    $xml = [xml](Get-Content C:\Windows\System32\inetsrv\config\applicationHost.config)

    ## Get central logging directory
    $central_LogDir = $xml.configuration."system.applicationHost".log.centralW3CLogFile.Directory

    if($central_LogDir -like "%SystemDrive%*")
    {
        $central_LogDir = $central_LogDir -replace "%SystemDrive%","C:"
    }

    $Dirs = Get-ChildItem $central_LogDir
    foreach ($logDirName in $Dirs)
    {
        $logDir = $central_LogDir + "\" + $logDirName
        
        $Files = [IO.Directory]::GetFiles($logDir)
        foreach($file in $Files)
        {
            $fileProperties = Get-ItemProperty -path $file
            $fileAge = New-TimeSpan -Start ($fileProperties.LastWriteTime) -end (Get-Date)
            
            if($fileAge.TotalDays -gt 30)
            {
                Remove-Item $file
            }
        }
    }   
}