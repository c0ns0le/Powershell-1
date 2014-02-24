CLEAR

$dateFull = Get-Date
$date = [string]$dateFull.Month + $dateFull.Day + $dateFull.Year

$direc = "\\rcbackup1\UES_SQL\" + $date
md $direc | out-null
Write-Host $date

#Copy-Item -Path \\oecfp3\SQLBU\OECEXTRANET\DEFAULT\FULL\ -Filter *.sqb -Destination \\rcbackup1\UES_SQL\ -Recurse