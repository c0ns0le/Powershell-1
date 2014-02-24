$ELSources = Get-EventLog -LogName "Application" | Select-Object Source -Unique | Where-Object {$_.Source -eq "Remote Process Kill"}

if ($ELSources -eq $null)
{
	New-EventLog -LogName "Application" -Source "Remote Process Kill"
}

try{
	##$CompName="W7DMYERSM1"
	##$ProcID=6048
	#$CompName="W7LWACLAWSKIJ"
	#$ProcID=6324
	
	$Process=Get-WMIObject -ComputerName $CompName -Query "SELECT * FROM win32_process WHERE ProcessId = '$ProcID'"
		
	$Process.Terminate()
	
	$PName=$Process.Name
	$msg = "============================= `n Target: $CompName `n ============================= `n Process Intended to Kill: $PName `n ============================= `n Completion Code: Success `n ============================="
	
	Write-EventLog -LogName "Application" -Source "Remote Process Kill" -EntryType Information -EventID 1611 -Message $msg
}
catch
{
	[system.exception]
	
	$msg = "============================= `n Process Intended to Kill: $Process.Name `n ============================= `n Completion Code: Failure `n ============================= `n Error(s): $Error `n ============================="
	
	Write-EventLog -LogName "Application" -Source "Remote Process Kill" -EntryType Information -EventID 1611 -Message $msg
}