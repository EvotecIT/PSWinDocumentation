function Start-WinDocumentationWorkstation {
    param (
        [string[]] $ComputerName = $Env:COMPUTERNAME,
        [string] $FilePath,
        [switch] $OpenDocument
    )

    $Data0 = Get-WmiObject win32_computersystem -ComputerName $ComputerName|select PSComputerName, Name, Manufacturer , Domain, Model , Systemtype, PrimaryOwnerName, PCSystemType, PartOfDomain, CurrentTimeZone, BootupState
    $Data1 = Get-WmiObject win32_bios -ComputerName $ComputerName| select Status, Version, PrimaryBIOS, Manufacturer, ReleaseDate, SerialNumber
    $Data2 = Get-WmiObject win32_DiskDrive -ComputerName $ComputerName | Select Index, Model, Caption, SerialNumber, Description, MediaType, FirmwareRevision, Partitions, @{Expression = {$_.Size / 1Gb -as [int]}; Label = "Total Size(GB)"}, PNPDeviceID
    $Data3 = get-WmiObject win32_networkadapter -ComputerName $ComputerName | Select Name, Manufacturer, Description , AdapterType, Speed, MACAddress, NetConnectionID, PNPDeviceID
    $Data4 = Get-WmiObject win32_startupCommand -ComputerName $ComputerName | select Name, Location, Command, User, caption
    $Data5 = Get-WmiObject win32_logicalDisk -ComputerName $ComputerName | select DeviceID, VolumeName, @{Expression = {$_.Size / 1Gb -as [int]}; Label = "Total Size(GB)"}, @{Expression = {$_.Freespace / 1Gb -as [int]}; Label = "Free Size (GB)"}
    $Data6 = get-WmiObject win32_operatingsystem -ComputerName $ComputerName | select Caption, Organization, InstallDate, OSArchitecture, Version, SerialNumber, BootDevice, WindowsDirectory, CountryCode
    $Data7 = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation | select Model, Manufacturer, Logo, SupportPhone, SupportURL, SupportHours
    $Data8 = get-culture | select KeyboardLayoutId, DisplayName, @{Expression = {$_.ThreeLetterWindowsLanguageName}; Label = "Windows Language"}

    $WordDocument = New-WordDocument $FilePath

    Add-WordText -WordDocument $WordDocument -Text 'Computer Name: ', "$ComputerName" -FontSize 10 -Bold $false, $true -ContinueFormatting
    Add-WordText -WordDocument $WordDocument -Text 'Run by:', "$ENV:USERNAME" -FontSize 10 -ContinueFormatting

    Add-WordText -WordDocument $WordDocument -Text 'Computer System' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data0 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Bios information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data1 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Disk Drive information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data2 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Netork Adaptor Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data3 -Design ColorfulGrid -AutoFit Window -MaximumColumns 10

    Add-WordText -WordDocument $WordDocument -Text 'Startup  Software Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data4 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Disk Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data5 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'OS Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data6 -Design ColorfulGrid -AutoFit Window

    if ($null -ne $Data7) {
        Add-WordText -WordDocument $WordDocument -Text 'OEM Information' -FontSize 10 -HeadingType Heading1
        Add-WordTable -WordDocument $WordDocument -DataTable $Data7 -Design ColorfulGrid -AutoFit Window
    }

    Add-WordText -WordDocument $WordDocument -Text 'Culture Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data8 -Design ColorfulGrid -AutoFit Window


    Save-WordDocument -WordDocument $WordDocument -Language 'en-US'
    if ($OpenDocument) { Invoke-Item $FilePath }
}