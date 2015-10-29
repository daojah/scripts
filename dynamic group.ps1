Import-Module ActiveDirectory
Set-Location ad:

Get-ADGroupMember "Syntegra Staff" | foreach {Remove-ADGroupMember "Syntegra Staff" $_ -Confirm:$false}

Get-ADUser -filter  {mail -like "*@syntegra.com.ua" -and enabled -eq $true } | foreach {Add-ADGroupMember "Syntegra Staff" -Member $_}


