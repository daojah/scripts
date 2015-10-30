$user = ""
$pwd = ""
$de = "LDAP://[SERVERNAME]/cn=user,ou=people,o=company"
$deObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry($de,$user,$pwd,'FastBind')
$query = new-Object System.DirectoryServices.DirectorySearcher($deObject)
#$query.searchroot = $root
$objClass = $query.findall() |
ForEach-Object 
{
Where-Object {$_.properties.objectclass -eq "contact" }
Write-Host ("Importing " + $_.Properties.name)

$val = $objClass
if ($val -eq $null)
{
write-Host "$strLine object does not exist"
}
else
{
Add-Content -Path C:\UAB\usersContacts.csv -Value $val -Encoding unicode
#$properties = $objClass.Properties
#write-Host "Name: $($properties.name)"
}
}

