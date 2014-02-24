Function Get-WindowsEdition
{
	$Server 		= Read-Host "What Server?"
	$ADInfo 		= Get-ADComputer -Filter {Name -Like $Server} -Property *
	
	$OSEdition		= $ADInfo.OperatingSystem
	$OSServicePack	= $ADInfo.OperatingSystemServicePack
	$OSVersion		= $ADInfo.OperatingSystemVersion
	
	"Windows Edition: " + $OSEdition + " " + $OSServicePack + ", Version: " + $OSVersion
	
}
Function Get-NetworkInfo
{
	$Server 		= Read-Host "What Server?"
	
	$ADInfo 		= Get-ADComputer -Filter {Name -Like $Server} -Property *
	$ComputerInfo	= Get-WmiObject win32_computersystem -ComputerName $Server
	$NetInfo		= Get-WmiObject win32_NetworkAdapterConfiguration -ComputerName $Server -Filter IPEnabled=TRUE
	
	$CName			= $ADInfo.CanonicalName
	$DNS			= $NetInfo.DNSServerSearchOrder
	$Domain			= $ComputerInfo.Domain
	$IPAddr			= $NetInfo.IPAddress
	$Gateway		= $NetInfo.DefaultIPGateway
	$Subnet			= $NetInfo.IPSubnet
	
	"=== Network Details ==="
	"Domain: `t" 			+ $Domain
	"Canonical Name: `t" 	+ $CName
	"IP Address(es): `t" 	+ $IPAddr
	"Subnet Mask: `t" 	+ $Subnet
	"Default Gateway: `t" + $Gateway
	"DNS Server(s): `t" 	+ $DNS
}