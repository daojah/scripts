Get-ADUser -LDAPFilter "(|(memberOf=cn=Staff,ou=Staff,dc=MyDomain,dc=com)(memberOf=cn=Staff,ou=Students,dc=MyDomain,dc=com))" -Properties sAMAccountName, GivenName, Surname, EmailAddress | Select sAMAccountName, GivenName, Surname, EmailAddress | Export-Csv Report.csv

Get-ADObject -filter 'objectclass -eq "user"' -properties * | select | export-csv -NoTypeInformation "c:\ADUsers.csv"