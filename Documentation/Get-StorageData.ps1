clear

$StorageInfo=Get-WmiObject win32_volume | Where-Object {$_.DriveType -eq 3 -and $_.Label -ne 'System Reserved' -and $_.DriveLetter -ne $null}

$DriveLetters=$StorageInfo | Select-Object DriveLetter | Sort-Object DriveLetter | ft -HideTableHeaders
$DriveNames=$StorageInfo | Sort-Object DriveLetter | Select-Object Label | ft -HideTableHeaders
$DriveCapacity=$StorageInfo |Sort-Object DriveLetter | ForEach-Object {[Math]::Truncate($_.Capacity / 1GB)}

$DriveLetters


$DriveNames


$DriveCapacity