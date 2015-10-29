Import-Module Azure #добавление модуля Azure для Power Shell

#Add-AzureAccount

Select-AzureSubscription "Smart subscription" # Выбор подписки

Get-AzureVM -ServiceName "Terminals" -Name "Terminal2" | Set-AzureVMSize –InstanceSize ExtraSmall | Update-AzureVM # Параметры Service name и Name, имя сервиса и имя виртуальной машины, можно узнать используя командлет "get-azurevm", параметр InstanceSize командлетом Get-AzureRoleSize | Select InstanceSize 