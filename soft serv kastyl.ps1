$exit=0
#$ol=0
while ($exit -eq 0)
{

Get-Mailbox andrey.bushmakin 
$ol=NewSearch-Mailbox -SearchQuery Subject:'RE: ZureView Query1' -Identity andrey.bushmakin -TargetMailbox ruslan.moldaliev@syntegra.com.ua -TargetFolder Inbox 
 
 if ( $ol.'ResultItemsCount' -EQ "0") 
 {
 Write-Host('eq=0') 
 sleep -seconds 20
 } 
 else {$exit = 1}  
}



 $ol.'ResultItemsCount' 

 