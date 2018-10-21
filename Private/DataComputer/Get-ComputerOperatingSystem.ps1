function Get-ComputerOperatingSystem {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data6 = get-WmiObject win32_operatingsystem -ComputerName $ComputerName | select Caption, Organization, InstallDate, OSArchitecture, Version, SerialNumber, BootDevice, WindowsDirectory, CountryCode
    return $Data6
}