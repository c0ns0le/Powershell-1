CLEAR

Import-Module ActiveDirectory
Import-Module DocumentationTools

$hostInputFile = "D:\Development\Powershell\Documentation\Documents\servers.txt"

$computers = Get-Content $hostInputFile

$msWord = New-Object -Com Word.Application

foreach ($computer in $computers)
{
	$docTemplate = "D:\Development\Powershell\Documentation\Documents\BLANK.doc"
	if (! (test-path $docTemplate))
	{
  		throw "$($docTemplate) is not a valid path"
	}
	
	$wordDoc = $msWord.Documents.Open($docTemplate)
 	$wordDoc.Activate()
 
 	Write-Host "/** Starting documentation for " $computer " **\"
	
	# Document Title
	
	
	# Save the document to disk and close it. Change $filename path to suit your environment.
	Write-Host "...Saving"
	$filename = "D:\Development\Powershell\Documentation\Output\" + $computer + ".doc"
	$wordDoc.SaveAs([REF]$filename)
	$wordDoc.Close()
	Write-Host "/* Saved! *\"
}

$msWord.Application.Quit()
Write-Host "/** Process completed **\"