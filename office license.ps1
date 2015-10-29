Import-Module MSOnline
$cred = Get-Credential
Connect-MsolService -Credential $cred
$licensedAccountinguserList=Get-MsolUser -All | where {$_.IsLicensed -eq "true" } ForEach-Object in $licensedAccountinguserList
{Set-MsolUserLicense -RemoveLicenses contoso:ENTERPRISEPACK –AddLicenses contoso:ENTERPRISEWITHSCAL 
     $O365Licences = New-MsolLicenseOptions –AccountSkuId contoso:ENTERPRISEWITHSCAL -DisabledPlans SHAREPOINTENTERPRISE, SHAREPOINTWAC}
                {Get-MsolUser -All | Set-MsolUserLicense -LicenseOptions $O365Licences}


 
    

    


 # (Get-MsolUser -UserPrincipalName `
                              
 # "JaneDoe@<tenant>.onmicrosoft.com").Licenses.ServiceStatus