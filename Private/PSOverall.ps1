function Get-ComputerData {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Data0 = Get-WmiObject win32_computersystem -ComputerName $ComputerName | select PSComputerName, Name, Manufacturer , Domain, Model , Systemtype, PrimaryOwnerName, PCSystemType, PartOfDomain, CurrentTimeZone, BootupState
    return $Data0
}
function Get-ComputerBios {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data1 = Get-WmiObject win32_bios -ComputerName $ComputerName| select Status, Version, PrimaryBIOS, Manufacturer, ReleaseDate, SerialNumber
    return $Data1
}
function Get-ComputerDisk {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Data2 = Get-WmiObject win32_DiskDrive -ComputerName $ComputerName | Select Index, Model, Caption, SerialNumber, Description, MediaType, FirmwareRevision, Partitions, @{Expression = {$_.Size / 1Gb -as [int]}; Label = "Total Size(GB)"}, PNPDeviceID
    return $Data2
}
function Get-ComputerDiskLogical {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data5 = Get-WmiObject win32_logicalDisk -ComputerName $ComputerName | select DeviceID, VolumeName, @{Expression = {$_.Size / 1Gb -as [int]}; Label = "Total Size(GB)"}, @{Expression = {$_.Freespace / 1Gb -as [int]}; Label = "Free Size (GB)"}
    return $Data5
}
function Get-ComputerNetwork {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data3 = get-WmiObject win32_networkadapter -ComputerName $ComputerName | Select Name, Manufacturer, Description , AdapterType, Speed, MACAddress, NetConnectionID, PNPDeviceID
    $Data3 = $Data3 | Select Name, Manufacturer, Speed, AdapterType, MACAddress
    return $Data3
}
function Get-ComputerStartup {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data4 = Get-WmiObject win32_startupCommand -ComputerName $ComputerName | select Name, Location, Command, User, caption
    $Data4 = $Data4 | Select Name, Command, User, Caption
    return $Data4
}
function Get-ComputerOperatingSystem {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data6 = get-WmiObject win32_operatingsystem -ComputerName $ComputerName | select Caption, Organization, InstallDate, OSArchitecture, Version, SerialNumber, BootDevice, WindowsDirectory, CountryCode
    return $Data6
}
function Get-ComputerOemInformation {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data7 = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation | select Model, Manufacturer, Logo, SupportPhone, SupportURL, SupportHours
    return $Data7
}
function Get-ComputerCulture {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $ScriptBlock = {
        get-culture | select KeyboardLayoutId, DisplayName, @{Expression = {$_.ThreeLetterWindowsLanguageName}; Label = "Windows Language"}
    }
    if ($ComputerName -eq $Env:COMPUTERNAME) {
        $Data8 = Invoke-Command -ScriptBlock $ScriptBlock
    } else {
        $Data8 = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock
    }
    return $Data8
}