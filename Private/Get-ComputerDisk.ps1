function Get-ComputerDisk {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Data2 = Get-WmiObject win32_DiskDrive -ComputerName $ComputerName | Select Index, Model, Caption, SerialNumber, Description, MediaType, FirmwareRevision, Partitions, @{Expression = {$_.Size / 1Gb -as [int]}; Label = "Total Size(GB)"}, PNPDeviceID
    return $Data2
}