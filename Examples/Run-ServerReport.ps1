
Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord

$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate-WordTemplate.docx"

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"

function Start-WinDocumentationServer {
    param (
        [string[]] $ComputerName = $Env:COMPUTERNAME,
        [string] $FilePathTemplate,
        [string] $FilePath,
        [switch] $OpenDocument
    )


    $WordDocument = Get-WordDocument $FilePathTemplate

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

    Add-WordText -WordDocument $WordDocument -Text 'Installed Applications' -FontSize 10 -HeadingType Heading1
    Add-WordTable -WordDocument $WordDocument -DataTable $Data10 -Design ColorfulGrid -AutoFit Window

    Save-WordDocument -WordDocument $WordDocument -Language 'en-US' -FilePath $FilePath
    if ($OpenDocument) { Invoke-Item $FilePath }
}


Start-WinDocumentationServer -ComputerName 'AD1' -FilePathTemplate $FilePathTemplate -FilePath