function Start-WinDocumentationWorkstation {
    param (
        [string[]] $ComputerName = $Env:COMPUTERNAME,
        [string] $FilePath,
        [switch] $OpenDocument
    )

    $Data0 = Get-ComputerData -ComputerName $ComputerName
    $Data1 = Get-ComputerBios -ComputerName $ComputerName
    $Data2 = Get-ComputerDisk -ComputerName $ComputerName
    $Data3 = Get-ComputerNetwork -ComputerName $ComputerName
    $Data4 = Get-ComputerStartup -ComputerName $ComputerName
    $Data5 = Get-ComputerDiskLogical -ComputerName $ComputerName
    $Data6 = Get-ComputerOperatingSystem -ComputerName $ComputerName
    $Data7 = Get-ComputerOemInformation -ComputerName $ComputerName
    $Data8 = Get-ComputerCulture -ComputerName $ComputerName
    $Data9 = Get-ComputerServices -ComputerName $ComputerName

    $WordDocument = New-WordDocument $FilePath

    Add-WordText -WordDocument $WordDocument -Text 'Computer Name: ', "$ComputerName" -FontSize 10 -Bold $false, $true -ContinueFormatting
    Add-WordText -WordDocument $WordDocument -Text 'Run by:', "$ENV:USERNAME" -FontSize 10 -ContinueFormatting

    Add-WordText -WordDocument $WordDocument -Text 'Computer System' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data0 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Bios information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data1 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Disk Drive information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data2 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Disk Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data5 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Netork Adaptor Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data3 -Design ColorfulGrid -AutoFit Window -MaximumColumns 10

    Add-WordText -WordDocument $WordDocument -Text 'Startup  Software Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data4 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'OS Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data6 -Design ColorfulGrid -AutoFit Window

    if ($null -ne $Data7) {
        Add-WordText -WordDocument $WordDocument -Text 'OEM Information' -FontSize 10 -HeadingType Heading1
        Add-WordTable -WordDocument $WordDocument -DataTable $Data7 -Design ColorfulGrid -AutoFit Window
    }

    Add-WordText -WordDocument $WordDocument -Text 'Culture Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data8 -Design ColorfulGrid -AutoFit Window

    Add-WordText -WordDocument $WordDocument -Text 'Services Information' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data9 -Design ColorfulGrid -AutoFit Window

    Save-WordDocument -WordDocument $WordDocument -Language 'en-US'
    if ($OpenDocument) { Invoke-Item $FilePath }
}