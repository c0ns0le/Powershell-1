# Writes data to Word using bookmarks

# If you want to specify credentials the WMI command, uncomment the $cred below, and also make sure to uncomment the -Credential $cred in the 2 functions.
#$cred = Get-Credential $User
CLEAR

Import-Module ActiveDirectory
Import-Module DocumentationTools

# specify full path to the list of servers and services you want to scan.
$hostInputFile = "D:\Development\Powershell\Documentation\Documents\servers_testing.txt"
$servicesFile = "D:\Development\Powershell\Documentation\Documents\services.txt"
$services = Get-Content $servicesFile

if (! (test-path $hostInputFile))
{
 throw "$($hostInputFile) is not a valid path."
}

$computers = Get-Content $hostInputFile

$msWord = New-Object -Com Word.Application

foreach ($computer in $computers)
{
# # Opening correct document
# if ($computer -like "*DB*")
# {
# 	# specify fill path to the word doc where bookmarks are defined.
#	$docTemplate = "D:\Development\Powershell\Documentation\Documents\BLANKdb.doc"
#	if (! (test-path $docTemplate))
#	{
#  		throw "$($docTemplate) is not a valid path"
#	}
# }
# elseif ($computer -like "*WEB*")
# {
# 	# specify fill path to the word doc where bookmarks are defined.
#	$docTemplate = "D:\Development\Powershell\Documentation\Documents\BLANKweb.doc"
#	if (! (test-path $docTemplate))
#	{
#  		throw "$($docTemplate) is not a valid path"
#	}
# }
# else
# {
 	# specify full path to the word doc where bookmarks are defined.
#	$docTemplate = "D:\Development\Powershell\Documentation\Documents\BLANK.doc"
#	if (! (test-path $docTemplate))
#	{
#  		throw "$($docTemplate) is not a valid path"
#	}
# }
 
 #$wordDoc = $msWord.Documents.Open($docTemplate)
 $wordDoc = $msWord.Documents.Open("D:\Development\Powershell\Documentation\Documents\BLANK.doc")
 $wordDoc.Activate()
 
 Write-Host "/** Starting documentation for " $computer " **\"
 
 # Document Title
 Write-Host "...Setting Title"
 $objRange = $wordDoc.Bookmarks.Item("docTitle").Range
 $objRange.Text = $computer
 $wordDoc.Bookmarks.Add("docTitle",$objRange) | Out-Null
 
 # Document Title_Footer
 Write-Host "...Setting Title on footer"
 $objRange = $wordDoc.Bookmarks.Item("docTitle_footer").Range
 $objRange.Text = $computer
 $wordDoc.Bookmarks.Add("docTitle_footer",$objRange) | Out-Null
 
 # Set Edit Date
 Write-Host "...Setting Edit Date"
 $objRange = $wordDoc.Bookmarks.Item("editDate").Range
 $objRange.Text = Get-Date -Format F
 $wordDoc.Bookmarks.Add("editDate",$objRange) | Out-Null
 
 # HOSTNAME input (from the $computer value above)
 Write-Host "...writing Hostname"
 $objRange = $wordDoc.Bookmarks.Item("hostname").Range
 $objRange.Text = $computer
 $wordDoc.Bookmarks.Add("hostname",$objRange) | Out-Null
 
 # FQDN input
 Write-Host "...writing Fully Qualified Domain Name"
 $objRange = $wordDoc.Bookmarks.Item("fqdn").Range
 $FullName = $computer + "." + (Get-Domain_OEC -CompName $computer)
 $objRange.Text = $FullName
 $wordDoc.Bookmarks.Add("fqdn",$objRange) | Out-Null
 
 # OS VERSION input
 Write-Host "...writing OS Version"
 $objRange = $wordDoc.Bookmarks.Item("osversion").Range
 $objRange.Text = Get-OSVersion_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("osversion",$objRange) | Out-Null

 # COMMISION DATE input
 Write-Host "...writing Commision Date"
 $objRange = $wordDoc.Bookmarks.Item("commision").Range
 $objRange.Text = Get-CommisionDate_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("commision",$objRange) | Out-Null
 
 # PROCESSOR input
 Write-Host "...writing Processor(s)"
 $objRange = $wordDoc.Bookmarks.Item("processor").Range
 $objRange.Text = Get-Processor_OEC -CompName $Computer
 $wordDoc.Bookmarks.Add("processor",$objRange) | Out-Null
 
 # MEMORY input
 Write-Host "...writing Memory"
 $objRange = $wordDoc.Bookmarks.Item("memory").Range
 $MemVar = Get-Memory_OEC -CompName $computer
 $objRange.Text = $MemVar
 $wordDoc.Bookmarks.Add("memory",$objRange) | Out-Null
 
 # STORAGE input
 Write-Host "...writing Storage"
 $func = Get-Storage_OEC -CompName $computer
 write-List_OEC -inputFunction $func -bookmarkName "storage"
 
 # DOMAIN input
 Write-Host "...writing Domain"
 $objRange = $wordDoc.Bookmarks.Item("domain").Range
 $objRange.Text = Get-Domain_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("domain",$objRange) | Out-Null
 
 # CANONICAL NAME input
 Write-Host "...writing Canonical Name"
 $objRange = $wordDoc.Bookmarks.Item("cname").Range
 #$CNameVar = canonicalName -CompName $computer
 $objRange.Text = Get-CanonicalName_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("cname",$objRange) | Out-Null
 
 # IP ADDRESS input
 Write-Host "...writing IP Address"
 $objRange = $wordDoc.Bookmarks.Item("ip").Range
 $objRange.Text = Get-IPAddr_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("ip",$objRange) | Out-Null
 
 # SUBNET input
 Write-Host "...writing Subnet Mask"
 $objRange = $wordDoc.Bookmarks.Item("subnet").Range
 $objRange.Text = Get-Subnet_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("subnet",$objRange) | Out-Null
 
 # GATEWAY input
 Write-Host "...writing Gateway"
 $objRange = $wordDoc.Bookmarks.Item("gateway").Range
 $objRange.Text = Get-Gateway_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("gateway",$objRange) | Out-Null
 
 # DNS input
 Write-Host "...writing Domain Name Servers"
 $objRange = $wordDoc.Bookmarks.Item("dns").Range
 $objRange.Text = Get-DNS_OEC -CompName $computer
 $wordDoc.Bookmarks.Add("dns",$objRange) | Out-Null
 
 # CRITICAL SERVICES - Names
 Write-Host "...writing Critical Services (NAMES)"
 $func = Get-CritServices_OEC -CompName $computer -reqData "names"
 write-List_OEC -inputFunction $func -bookmarkName "services"
 
 # CRITICAL SERVICES - Run-As Accounts
 Write-Host "...writing Critical Services (RUNAS)"
 $func = Get-CritServices_OEC -CompName $computer -reqData "runas"
 write-List_OEC -inputFunction $func -bookmarkName "services_runas"
 
 # CRITICAL SERVICES - Startup Type
 Write-Host "...writing Critical Services (STARTUP_TYPE)"
 $func = Get-CritServices_OEC -CompName $computer -reqData "start"
 write-List_OEC -inputFunction $func -bookmarkName "services_start"
 
 # INSTALLED SOFTWARE
 Write-Host "...writing Installed Software"
 $func = Get-InstalledSW_OEC -CompName $computer -reqData "names"
 write-List_OEC -inputFunction $func -bookmarkName "installedsw"
 
 # INSTALLED SOFTWARE - Version
 Write-Host "...writing Installed Software"
 $func = Get-InstalledSW_OEC -CompName $computer -reqData "version"
 write-List_OEC -inputFunction $func -bookmarkName "installedsw_version"
 
# if ($computer -like "*DB*")
# {
# 	# SQL INSTANCE input
#	Write-Host "...writing SQL Instance"
# 	$objRange = $wordDoc.Bookmarks.Item("sql_instance").Range
# 	$objRange.Text = Get-SQLinfo_OEC -CompName $computer -OpRequest 'Instance'
# 	$wordDoc.Bookmarks.Add("sql_instance",$objRange) | Out-Null
#	
#	# SQL EDITION input
#	Write-Host "...writing SQL Edition"
# 	$objRange = $wordDoc.Bookmarks.Item("sql_edition").Range
# 	$objRange.Text = Get-SQLinfo_OEC -CompName $computer -OpRequest 'Edition'
# 	$wordDoc.Bookmarks.Add("sql_edition",$objRange) | Out-Null
#	
#	# SQL LOGIN METHOD(S) input
#	Write-Host "...writing SQL Login Method"
# 	$objRange = $wordDoc.Bookmarks.Item("sql_login").Range
# 	$objRange.Text = Get-SQLinfo_OEC -CompName $computer -OpRequest 'Login'
# 	$wordDoc.Bookmarks.Add("sql_login",$objRange) | Out-Null
#	
#	# SQL ADMIN(S) input
#	Write-Host "...writing SQL SysAdmins"
#	$func = Get-SQLinfo_OEC -CompName $computer -OpRequest 'Admins'
# 	write-List_OEC -inputFunction $func -bookmarkName "sql_admins"
# }
# elseif ($computer -like "*WEB*")
# {
# 	Write-Host "Web Stuff"
# }
# else
# {
# 	Write-Host "Unknown Stuff"
# }
 
 # Save the document to disk and close it. Change $filename path to suit your environment.
 Write-Host "...Saving"
 $filename = "D:\Development\Powershell\Documentation\Output\" + $computer + ".doc"
 $wordDoc.SaveAs([REF]$filename)
 $wordDoc.Close()
 Write-Host "/* Saved! *\"
}

$msWord.Application.Quit()
Write-Host "/** Process completed **\"