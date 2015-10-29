$user = "username"
$pwd = "password"
$de = "LDAP://[SERVERNAME]/cn=user,ou=people,o=company"
$deObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry($de,$user,$pwd,'FastBind')
$selector = new-Object System.DirectoryServices.DirectorySearcher($deObject)
$Selector.SearchRoot = $de
$selector.Filter = '(objectClass=Contact)'
$selector.FindAll() | Foreach 
{
Write-Host ("Importing " + $_.Properties.name)
select @{name="Имя";expression={$_.Properties.name}},
@{n="EMail";E={$_.Properties.mail}} >C:\1.csv
}



