Param($user,
      $password = $(Read-Host "Enter Password" -asSec),
      $filter = "(objectclass=user)",
      $server = $(throw  '$server is required'),    #тут адрес сервака “server.cgp.local:489”    
      $path = $(throw '$path is required'),         #тут путь додлжен быть что-то типо “cn=cgp.local,o=mxs”
      [switch]$all,
      [switch]$verbose)
    
function GetSecurePass ($SecurePassword) {
  $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($SecurePassword)
  $password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
  [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
  $password
}    

if($verbose){$verbosepreference = "Continue"}

$DN = "LDAP://$server/$path"
Write-Verbose "DN = $DN"
$auth = [System.DirectoryServices.AuthenticationTypes]::FastBind
Write-Verbose "Auth = FastBind"
$de = New-Object System.DirectoryServices.DirectoryEntry($DN,$user,(GetSecurePass $Password),$auth)
Write-Verbose $de
Write-Verbose "Filter: $filter"
$ds = New-Object system.DirectoryServices.DirectorySearcher($de,$filter) 
Write-Verbose $ds

Write-Verbose "Finding All"
$ds.FindAll() >C:/1.txt

$Encode =  new-object system.text.UTF8encoding
Set-Content -Path C:\UsersContacts.csv -Value “displayname, alias, WindowsEmailAddress” -Encoding unicode
$i = 0
foreach ($user in $users)
{
    $i++; Write-Host $i
    $Row = “{0}, {1}, {2}” -f $Encode.GetString($user.properties[“cn”][0]), $Encode.GetString($user.properties[“uid”][0]), $Encode.GetString($user.properties[“mail”][0])
    Add-Content -Path C:\usersContacts.csv -Value $Row -Encoding unicode
}
