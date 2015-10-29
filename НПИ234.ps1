$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)


$inbox.items | foreach {
        if($_.subject -match "RE: ZureView Query1"){ls}
        
}