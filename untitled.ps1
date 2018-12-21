Get-Folder -Type VM | %{
    $_.Name
    $_
    }
	
#connect-viserver -server "xxx" -user "xxx" -password "xxx"
$report = @()
$vms = Get-View -ViewType "virtualmachine"
foreach($vm in $vms){
     foreach($dev in $vm.Config.Hardware.Device){
          if(($dev.gettype()).Name -eq "VirtualDisk"){
                    $row = "" | select VMName, FolderPath, PowerState, HostName, ClusterName, OSVersion, GuestIPAddress, ToolsStatus, NumCPU, MemoryMB, DeviceInfo, CapacityInGB, Datastore
   $row.VMName = $vm.Name
   $current = Get-View $vm.Parent
   $path = $vm.Name
   do {
     $parent = $current
     if($parent.Name -ne "vm"){$path =  $parent.Name + "\" + $path}
     $current = Get-View $current.Parent
   } while ($current.Parent -ne $null)
   $row.FolderPath = $path
   $row.PowerState = $vm.runtime.powerstate
   $row.HostName = (get-view -id $vm.runtime.host).name
   $clu = (get-view -id $vm.Runtime.Host).parent
   $row.ClusterName = (Get-View -id $clu).name
   $row.OSVersion = $vm.summary.guest.guestfullname
   $row.GuestIPAddress = $vm.guest.ipaddress
   $row.ToolsStatus = $vm.guest.toolsstatus
   $row.NumCPU = $vm.config.hardware.numcpu
   $row.MemoryMB = $vm.config.hardware.MemoryMB
   $row.DeviceInfo = $dev.deviceinfo.label
   $row.CapacityInGB = [system.math]::Round($dev.CapacityInKB / 1048576)
   $row.Datastore = $dev.backing.filename.split("]")[0].trim("[")
   $report += $row
        }
    }
}
$report | sort clusterName | ft #| Export-Csv C:\exportvms.csv -NoTypeInformation -UseCulture

