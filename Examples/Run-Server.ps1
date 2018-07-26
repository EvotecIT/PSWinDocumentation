Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord -Force
Import-Module ActiveDirectory

$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate.docx"

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"

function Start-WinDocumentationServer {
    param (
        [string[]] $ComputerName = $Env:COMPUTERNAME,
        [string] $FilePathTemplate,
        [string] $FilePath,
        [switch] $OpenDocument,
        $CompanyName = 'Evotec'
    )

    $WordDocument = New-WordDocument -FilePath $FilePath

    $Domain = 'ad.evotec.xyz'
    #$Domains = (Get-ADForest).Domains
    # foreach ($Domain in $Domains) {
    $ADSnapshot = Get-ActiveDirectoryCleanData -Domain $Domain
    $ActiveDirectorySnapshot = Get-ActiveDirectoryProcessedData -ADSnapshot $ADSnapshot
    #}
    $Toc = Add-WordToc -WordDocument $WordDocument -Title 'Table of content' -Switches C, A -RightTabPos 15 -HeaderStyle Heading1

    ### 1st section - Introduction
    $SectionScope = $WordDocument | Add-WordTocItem -Text 'Scope' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionScope = $WordDocument | Get-DocumentScope -Paragraph $SectionScope -ActiveDirectorySnapshot $ActiveDirectorySnapshot -CompanyName $CompanyName
    ### 2nd section - General Information
    # Leaving empty for now
    $SectionGeneralInformation = $WordDocument | Add-WordTocItem -Text 'General Information' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    ### 3rd section - Forest Summary
    $SectionForestSummary = $WordDocument | Add-WordTocItem -Text 'General Information - Forest Summary' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionForestSummary = $WordDocument | Get-ForestSummary -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ActiveDirectorySnapshot
    $SectionForestSummary = $WordDocument | Get-ForestFeatures -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ActiveDirectorySnapshot
    $SectionForestSummary = $WordDocument | Get-ForestFSMORoles -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ActiveDirectorySnapshot
    ### Section 4 - Domain Summary
    $SectionDomainSummary = $WordDocument | Add-WordTocItem -Text 'General Information - Domain Summary' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionDomainSummary = $WordDocument | Get-DomainSummary -Paragraph $SectionDomainSummary -ActiveDirectorySnapshot $ActiveDirectorySnapshot
    ### Section 5 - UPN Summary
    $SectionDomainUPNs = $WordDocument | Add-WordTocItem -Text 'General Information - UPN Summary' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionDomainUPNs = $WordDocument | Get-DomainUPNSuffixes -Paragraph $SectionDomainUPNs -ActiveDirectorySnapshot $ActiveDirectorySnapshot
    ### Section 6 - Password Policies
    $SectionPasswordPolicies = $WordDocument | Add-WordTocItem -Text 'General Information - Password Policies' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionPasswordPolicies = $WordDocument | Get-DomainPasswordPolicies -Paragraph $SectionPasswordPolicies -ActiveDirectorySnapshot $ActiveDirectorySnapshot

    Save-WordDocument -WordDocument $WordDocument -Language 'en-US' -FilePath $FilePath -Supress $true
    if ($OpenDocument) { Invoke-Item $FilePath }
}

Start-WinDocumentationServer -ComputerName 'AD1' -FilePathTemplate $FilePathTemplate -FilePath $FilePath -OpenDocument