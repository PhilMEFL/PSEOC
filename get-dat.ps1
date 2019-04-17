param (
 [parameter(mandatory=$true)][string]$computername
)
function ping-host([string]$computername) {
 #This function will perform a simple, small size single packet ping of a machine and return true/false for the result
  if ([string]::IsNullOrEmpty($computername) ) {return $false}
  #ping first for reachability check
  $po = New-Object net.NetworkInformation.PingOptions
  $po.set_ttl(64)
  $po.set_dontfragment($true)
  [Byte[]] $pingbytes = (65,72,79,89)
  $ping = new-object Net.NetworkInformation.Ping
  $savedEA = $Erroractionpreference
  $ErrorActionPreference = "silentlycontinue"
  $pingres = $ping.send($computername, 1000, $pingbytes, $po)
  if (-not $?) {return $false}
  $ErrorActionPreference = $savedEA
  if ($pingres.status -eq "Success") { return $true } else {return $false}
}


if ((ping-host $computername) -eq $false) {
 New-Object PSobject -Property @{
  Computername = $computername
  DATVersion = "System Not Online"
  Datdate = $null
 }
} else {

 try {
  #Set up the key that needs to be accessed and what registry tree it is under
  $key = "Software\McAfee\AVEngine"
  $type = [Microsoft.Win32.RegistryHive]::LocalMachine

  #open up the registry on the remote machine and read out the TOE related registry values
  $regkey = [Microsoft.win32.registrykey]::OpenRemoteBaseKey($type,$computername)
  $regkey = $regkey.opensubkey($key)
  $status = $regkey.getvalue("AVDatVersion")
  $datdate = $regkey.getvalue("AVDatDate")
 } catch {
  try {
   $key = "Software\Wow6432Node\McAfee\AVEngine"
   $type = [Microsoft.Win32.RegistryHive]::LocalMachine
   #open up the registry on the remote machine and read out the TOE related registry values
   $regkey = [Microsoft.win32.registrykey]::OpenRemoteBaseKey($type,$computername)
   $regkey = $regkey.opensubkey($key)
   $status = $regkey.getvalue("AVDatVersion")
   $datdate = $regkey.getvalue("AVDatDate")
  } catch {
     #try newer registry location
    try {
     $key = "Software\Wow5432Node\Network Associates\ePolicy Orchestrator\Application Plugins\VIRUSCAN880"
     $regkey = [Microsoft.win32.registrykey]::OpenRemoteBaseKey($type,$computername)
     $regkey = $regkey.opensubkey($key)
     $status = $regkey.getvalue("DATVersion")
     $datdate = $regkey.getvalue("DatDate")
    } catch {
      $status = "Cannot read regkey"
    }
  }
 }
 New-Object PSobject -Property @{
  Computername = $computername
  DATVersion = $status
  DatDate = $datdate
 } |select Computername,DatVersion,DatDate
}