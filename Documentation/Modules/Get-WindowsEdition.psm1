Function Get-WindowsEdition
{
	$Server 		= Read-Host "What Server?"
	$ADInfo 		= Get-ADComputer -Filter {Name -Like $Server} -Property *
	
	$OSEdition		= $ADInfo.OperationSystem
	$OSServicePack	= $ADInfo.OperatingSystemServicePack
	$OSVersion		= $ADInfo.OperatingSystemVersion
	
	"Windows Edition: " + $OSEdition + " " + $OSServicePack + ", Version: " + $OSVersion
	
}