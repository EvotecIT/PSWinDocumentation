function Get-ComputerBios {
    [CmdletBinding()]
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data1 = Get-WmiObject win32_bios -ComputerName $ComputerName| select Status, Version, PrimaryBIOS, Manufacturer, ReleaseDate, SerialNumber
    return $Data1
}