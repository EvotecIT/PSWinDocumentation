function Get-ComputerServices {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Services = Get-Service -ComputerName $ComputerName | select Name, Displayname, Status
    return $Services
}