Connect-VIServer -Server copvc01
get-datacenter "COP-VMwareDatacenter" | get-vm | get-networkadapter | where { $_.Type -ne 'Vmxnet3'}  | select Parent
#| set-networkadapter -networkname "NewPortGroup" -Confirm:$false