function Get-ComputerMissingDrivers {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Data = Get-WmiObject Win32_PNPEntity -ComputerName $ComputerName | where {$_.Configmanagererrorcode -ne 0} | Select Caption, ConfigmanagererrorCode, Description, DeviceId, HardwareId, PNPDeviceID
    return $Data
}
