Get-ADUser -filter { mail -like "*" } | foreach
{Add-ADGroupMember "1" -Member $_}