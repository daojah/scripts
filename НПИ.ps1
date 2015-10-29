$user = "abadmin"
$pwd = "1qazxsw2"
$de = "LDAP://10.0.0.3/OU=it,dc=syntegra,dc=ua"
$deObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry($de,$user,$pwd,'FastBind')
$selector = new-Object System.DirectoryServices.DirectorySearcher($deObject)
$Selector.SearchRoot = $de
$selector.Filter = '(objectclass=user)' 
$selector.Filter = '(mail=*)' 

$selector.PropertiesToLoad.AddRange(@("name"))
$selector.PropertiesToLoad.AddRange(@("mail"))
$users = $selector.FindAll()   
#where {$users.Propertiesloaded  -contains  "mail"}  
$users.Count
$report = @()
foreach ($objResult in $users)
{$objItem = $objResult.Properties
$temp = New-Object PSObject
$temp | Add-Member NoteProperty name $($objitem.name) -Force
$temp | Add-Member NoteProperty mail $($objitem.mail) -Force
$report += $temp
}
$report | export-csv -Path C:\1.csv -Encoding unicode
"Done!"