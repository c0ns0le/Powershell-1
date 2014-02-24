CLEAR

function Get-FolderSize_SA
{
	param([string]$dir)
	$rawSize = Get-ChildItem $dir -recurse | Measure-Object -property length -sum	
	
	"{0:N2}" -f ($rawSize.sum / 1GB) + " GB"
}

$TotalSize = "{0:N2}" -f (((Get-WmiObject win32_logicalDisk -ComputerName RCBACKUP1 | Where-Object {$_.DeviceID -eq "D:"}).Size) / 1GB)

$CRM_size = Get-FolderSize_SA -dir "\\rcbackup1\CRM_SQL"
	$CRMSize_Dec = [decimal]($CRM_size -replace " GB","")
$NAV_size = Get-FolderSize_SA -dir "\\rcbackup1\NAV_SQL"
	$NAVSize_Dec = [decimal]($NAV_size -replace " GB","")
$RPRT_size = Get-FolderSize_SA -dir "\\rcbackup1\RPRT_SQL"
	$RPRTSize_Dec = [decimal]($RPRT_size -replace " GB","")
$WSB_size = Get-FolderSize_SA -dir "\\rcbackup1\WSB_SQL"
	$WSBSize_Dec = [decimal]($WSB_size -replace " GB","")
$SCOM_size = Get-FolderSize_SA -dir "\\rcbackup1\SCOM_SQL"
	$SCOMSize_Dec = [decimal]($SCOM_size -replace " GB","")
$UES_size = Get-FolderSize_SA -dir "\\rcbackup1\UES_SQL"
	$UESSize_Dec = [decimal]($UES_size -replace " GB","")
$GES_size = Get-FolderSize_SA -dir "\\rcbackup1\GES_Logs"
	$GESSize_Dec = [decimal]($GES_size -replace " GB","")
	
$TotalUsed = $CRMSize_Dec + $NAVSize_Dec + $RPRTSize_Dec + $WSBSize_Dec + $SCOMSize_Dec + $UESSize_Dec + $GESSize_Dec

$msgBody =  Write-Output "Folder sizes for RCBACKUP1" `
						"`n>> CRM Directory: " $CRM_size `
						"`n>> NAV Directory: " $NAV_size `
                        "`n>> RPRT Directory: " $RPRT_size `
						"`n>> WSB Directory: " $WSB_size `
						"`n>> SCOM Directory: " $SCOM_size `
						"`n>> UES Directory: " $UES_size `
                        "`n>> GES Directory: " $GES_size `
						"`n" `
						"`n>> Total Space Used: " $TotalUsed "GB of" $TotalSize "GB" 

#Write-Host $msgbody

$smtpServer = "10.1.60.10"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

$msg.From = "support@oeconnection.com"
$msg.To.Add("Josh.Waclawski@oeconnection.com")
$msg.Subject = "SQL Backups Drive Space Usage"
$msg.Body = $msgBody
$smtp.Send($msg)
