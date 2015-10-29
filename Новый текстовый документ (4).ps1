
$user="DemoUser"
$pwd="Qwerty12"

$location="North Europe" 
$img="a699494373c04fc0bc8f2bb1389d6106__Windows-Server-2012-R2-201411.01-en.us-127GB.vhd"

$size="Small"

$cloudservice="ololololo123"

$vmConfig = New-AzureVMConfig -Name $cloudService -ImageName $img -InstanceSize $size



$vmConfig | Add-AzureProvisioningConfig -Windows -AdminUsername $user -Password $pwd  



$vmConfig | Add-AzureDataDisk -CreateNew  -DiskSizeInGB 100 -DiskLabel "Data 1" -Lun 0 


$vmConfig | Add-AzureDataDisk -CreateNew  -DiskSizeInGB 100 -DiskLabel "Data 2" -Lun 1 


$vmConfig | Add-AzureEndpoint -Name "Web" -Protocol tcp -LocalPort 80 -PublicPort 80 

 
