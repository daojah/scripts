$user = ""
$pwd = ""
$de = "LDAP://10.0.0.3/OU=it,dc=syntegra,dc=ua"
$deObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry($de,$user,$pwd,'FastBind')
$selector = new-Object System.DirectoryServices.DirectorySearcher($deObject)
$Selector.SearchRoot = $de
$selector.Filter = '(objectClass=user)' 
$users = $selector.FindAll()  
#foreach ($users in $users) 
{
select @{name="Имя";expression={$_.Properties.name}},
@{n="EMail";E={$_.Properties.mail}} >C:\1.csv
}
