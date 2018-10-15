function Get-ComputerData {
    [CmdletBinding()]
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Data0 = Get-WmiObject win32_computersystem -ComputerName $ComputerName | select PSComputerName, Name, Manufacturer , Domain, Model , Systemtype, PrimaryOwnerName, PCSystemType, PartOfDomain, CurrentTimeZone, BootupState
    return $Data0
}