Import-Module Azure 

Add-AzureAccount

Select-AzureSubscription "Smart subscription"

Get-AzureVM -ServiceName "Terminals" -Name "Terminal1" `| Set-AzureVMSize –InstanceSize "A6" `| Update-AzureVM