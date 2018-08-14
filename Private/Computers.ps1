function Get-ComputerData {
    [CmdletBinding()]
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Data0 = Get-WmiObject win32_computersystem -ComputerName $ComputerName | select PSComputerName, Name, Manufacturer , Domain, Model , Systemtype, PrimaryOwnerName, PCSystemType, PartOfDomain, CurrentTimeZone, BootupState
    return $Data0
}
function Get-ComputerBios {
    [CmdletBinding()]
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
    $ScriptBlock = { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation | select Model, Manufacturer, Logo, SupportPhone, SupportURL, SupportHours }
    if ($ComputerName -eq $Env:COMPUTERNAME) {
        $Data = Invoke-Command -ScriptBlock $ScriptBlock
    } else {
        $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock
    }
    return $Data
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
function Get-ComputerServices {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Services = Get-Service -ComputerName $ComputerName | select Name, Displayname, Status
    return $Services
}
function Get-ComputerApplications {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $ScriptBlock = {
        $objapp1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*
        $objapp2 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*

        $app1 = $objapp1 | Select-Object Displayname, Displayversion , Publisher, Installdate, @{Expression = { 'x64' }; Label = "WindowsType"}
        $app2 = $objapp2 | Select-Object Displayname, Displayversion , Publisher, Installdate, @{Expression = { 'x86' }; Label = "WindowsType"} | where { -NOT (([string]$_.displayname).contains("Security Update for Microsoft") -or ([string]$_.displayname).contains("Update for Microsoft"))}
        $app = $app1 + $app2 #| Sort-Object -Unique
        return $app | Where { $_.Displayname -ne $null } | Sort-Object DisplayName
    }
    if ($ComputerName -eq $Env:COMPUTERNAME) {
        $Data = Invoke-Command -ScriptBlock $ScriptBlock
    } else {
        $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock
    }
    return $Data

}

function Get-ComputerWindowsFeatures {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data = Get-WmiObject Win32_OptionalFeature -ComputerName $vComputerName | select Caption , Installstate
    return $Data
}

function Get-ComputerWindowsUpdates {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data = Get-hotfix -ComputerName $vComputerName | select Description , HotFixId , InstalledBy, InstalledOn, Caption
    return $Data
}



function Get-ComputerMissingDrivers {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $Data = Get-WmiObject Win32_PNPEntity -ComputerName $ComputerName | where {$_.Configmanagererrorcode -ne 0} | Select Caption, ConfigmanagererrorCode, Description, DeviceId, HardwareId, PNPDeviceID
    return $Data
}
