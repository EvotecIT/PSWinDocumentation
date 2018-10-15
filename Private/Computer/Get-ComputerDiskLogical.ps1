function Get-ComputerDiskLogical {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data5 = Get-WmiObject win32_logicalDisk -ComputerName $ComputerName | select DeviceID, VolumeName, @{Expression = {$_.Size / 1Gb -as [int]}; Label = "Total Size(GB)"}, @{Expression = {$_.Freespace / 1Gb -as [int]}; Label = "Free Size (GB)"}
    return $Data5
}