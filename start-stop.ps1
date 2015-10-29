function stopvm($azurecloud,$vmname) 
{ 
    if ($vmname -eq $null) 
    { 
        Get-AzureVM -ServiceName $azurecloud |Foreach-object {Stop-AzureVM -ServiceName $_.ServiceName -Name $_.Name -Force} 
    } 
    else 
    { 
        Stop-AzureVM -ServiceName $azurecloud -Name $vmname 
    } 
} 
function startvm($azurecloud,$vmname) 
{ 
    if ($vmname -eq $null) 
    { 
        Get-AzureVM -ServiceName $azurecloud |Foreach-object {Start-AzureVM -ServiceName $_.ServiceName -Name $_.Name -Force} 
    } 
    else 
    { 
        Start-AzureVM -ServiceName $azurecloud -Name $vmname 
    } 
}

################################################################ 
# Please Change These Variables to Suit Your Environment 
# 
$azuresettings = "C:\Azure\AzureseetingsFile" 
$azurecloud = "CloudServiceName" 
# 
# 
# 
################################################################

write-host "Importing Azure Settings" 
Import-AzurePublishSettingsFile $azuresettings

write-host "Choose the options to Start and Stop your Azure VMS" 
write-host "1. Start All VMs" 
write-host "2. Stop All VMs" 
write-host "3. Start One VM" 
write-host "4. Stop One VM" 
$answer = read-host "Please Select Your Choice"

Switch($answer) 
{ 
    1{ $vmname = $null;StartVM $azurecloud $vmname} 
    2{ $vmname = $null;StopVM $azurecloud $vmname} 
    3{ $vmname = read-host "Please Enter VM Name";StartVM $azurecloud $vmname} 
    4{ $vmname = read-host "Please Enter VM Name";StopVM $azurecloud $vmname} 
}