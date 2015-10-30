$xx =1
$user = ""
$pwd = ""
$de = "LDAP://10.0.0.3/ou=it,dc=syntegra,dc=ua"
$deObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry($de,$user,$pwd,'FastBind')
$selector = new-Object System.DirectoryServices.DirectorySearcher($deObject)
$Selector.SearchRoot = $de
$selector.Filter = '(objectClass=user)' 
$selector.PropertiesToLoad.AddRange(@("name"))
$users = $selector.Findall() |  `
where {$_.Properties.objectcategory -match  "mail"}


$users.Count  
$report = @()
foreach ($objResult in $users)  
<#{
If ($users.'mail' -ne  "")
 #>
{
$objItem = $objResult.Properties
$temp = New-Object PSObject
$temp | Add-Member NoteProperty name $($objitem.name) -Force
$temp | Add-Member NoteProperty mail $($objitem.mail) -Force
$report += $temp
$xx = $xx+1
}

$report | export-csv -Path C:\1.csv -Encoding unicode  -Delimiter "," -Force  
"Done!"


#foreach ($objResult in $users)  {Where-Object $report.'mail' -eq  ""} 
#{
#If ($users.'mail' -eq  "" )
 #{Write-Host "No email"}
#Else}
#{$objItem = $objResult.Properties
#$temp = New-Object PSObject
#$temp | Add-Member NoteProperty name $($objitem.name) -Force
#$temp | Add-Member NoteProperty mail $($objitem.mail) -Force
#$report += $temp
#}
#$report | export-csv -Path C:\1.csv -Encoding unicode  -Delimiter "," -Force  
#"Done!"

