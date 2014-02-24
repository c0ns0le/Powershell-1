CLEAR

# Build the function to GET storage data
# Parameter(s):
# -CompName <computer_name>
function storage
{
 param([string]$CompName)
 Get-WmiObject win32_volume -ComputerName $CompName -Filter "DriveType=3 AND Label <> 'System Reserved' AND DriveLetter IS NOT NULL" | Sort DriveLetter | ForEach-Object{"{0}, {1} - {2}gb" -f $_.Name,$_.Label,([Math]::Truncate($_.Capacity/1GB))}
}

function installedSW
{
 param([string]$CompName)
 Get-WmiObject win32_Product -ComputerName $CompName | Sort Name
}

# Create Microsoft Word ComObject
$msWord = New-Object -Com Word.Application

# Check path for validity; error out if the file does not exist
$docTemplate = "C:\Users\waclawskij\Desktop\Scripts\Documentation\Documents\BLANK.doc"
if (! (test-path $docTemplate))
{
  throw "$($docTemplate) is not a valid path"
}

# Open template and set as active document
Write-Host "== STARTING =="
$wordDoc = $msWord.Documents.Open($docTemplate)
$wordDoc.Activate()

# Build array of storage values; count total objects in array
Write-Host "...building array"
$objArray = InstalledSW -CompName 'RCSDVDB1'
$objArrayCount = $objArray.Count

# Locate the 'storage' Bookmark within the Word document
Write-Host "...building range objects"
$ObjRangeName = $wordDoc.Bookmarks.Item("installedsw").Range
$ObjRangeVersion = $wordDoc.Bookmarks.Item("installedsw_version").Range

# Add a row for every object within $storageArray
# This is designed to add one less row than there are drives on the server.
# This is done so because the template starts with one row already available
Write-Host "...adding rows"
for ($i=1; $i -lt ($objArrayCount-1); $i++)
{
 $Rows = $ObjRangeName.Rows.Add()
 Write-Host ".........row " $i " added"
}

# Set table coordinates of storage bookmark
Write-Host "...pinpointing coordinates"
$storageXname = $ObjRangeName.Information(14) # Gets the ROW number for the storage Bookmark
$storageYname = $ObjRangeName.Information(17) # Gets the COL number for the storage Bookmark
$storageXversion = $ObjRangeVersion.Information(14) # Gets the ROW number for the storage Bookmark
$storageYversion = $ObjRangeVersion.Information(17) # Gets the COL number for the storage Bookmark

# Add each drive to the document, starting at ($storageX,$storageY) and incrementing as needed
Write-Host "...adding values"
foreach ($object in $objArray)
{
  $Table = $wordDoc.Tables.Item(1)
  $tableRangeName = $Table.Cell($storageXname,$storageYname).Range.Text = $object.Name
  $tableRangeName = $Table.Cell($storageXversion,$storageYversion).Range.Text = $object.Version
  Write-Host "........." $object.Name " added"
  $storageXname++
  $storageXversion++
}

# Save the file as testing.doc
Write-Host "...Saving"
$filename = "C:\Users\waclawskij\Desktop\Scripts\Documentation\Output\testing.doc"
$wordDoc.SaveAs([REF]$filename)
$wordDoc.Close()
Write-Host "/* Saved! *\"

#CLEAR
#
#function Get-InstalledSW_OEC
#{
# param([string]$CompName,
# 		[string]$OutRequest)
# $Var1 = Get-WmiObject win32_Product -ComputerName $CompName | Sort Name #| ForEach-Object{"{0}" -f $_.Name}
# 
# $Var2 = $Var1 | ForEach-Object {"{0}" -f $_.Name}
# $Var3 = $Var1 | ForEach-Object {"{0}" -f $_.Version}
# 
# if($OutRequest = "Name")
# {
#  return $Var2
# }
# elseif($OutRequest = "Version")
# {
#  return $Var3
# }
# else
# {
#  Write-Host "Nope"	
# }
#}
#
#Get-InstalledSW_OEC -CompName "RCSDVDB1" -OutRequest "Version"