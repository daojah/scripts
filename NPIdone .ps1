$user = "Vadim.martynov"
$pwd = "Rjycfknbu01"
$de = "LDAP://dc2.ad.novaposhta.international/ou=NPI,dc=ad,dc=novaposhta,dc=international"
$deObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry($de,$user,$pwd,'FastBind')
$selector = new-Object System.DirectoryServices.DirectorySearcher($deObject)
$Selector.SearchRoot = $de
#$selector.Filter = '(objectclass=user)' 
#$selector.Filter = '(mail=*)' 
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
 $report | sort name | export-csv -Path C:\Users\baron\Desktop\1.csv -Encoding unicode -Delimiter "," -Force 
"Done!"

Start-Sleep -Seconds 1800

$PWord = ConvertTo-SecureString –String $pwd –AsPlainText -Force
$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $user, $PWord
connect-MSOLService -Credential $Credential

#Задаем переменные
$CSVpatch = "C:\Users\baron\Desktop\1.csv"
Import-Csv $CSVpatch  | 
 foreach-object if (Get-MailContact -anr $_.name) {write-host $_.name 'is a duplicate entry!!!'}  
  else   {New-MailContact -Name $_.Name -DisplayName $_.Name -ExternalEmailAddress $_.Mail }


