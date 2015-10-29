Import-Module Azure 

Add-AzureAccount

Select-AzureSubscription "Smart subscription"

Get-AzureVM -ServiceName "Terminals" -Name "Terminal2" `| Set-AzureVMSize –InstanceSize Medium `| Update-AzureVM