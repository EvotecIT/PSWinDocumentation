Import-Module PSWInDocumentation #-Force
Import-Module PSWriteWord # -Force
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
    if ($FilePathTemplate -ne $null) {
        $WordDocument = Get-WordDocument -FilePath $FilePathTemplate
    } else {
        $WordDocument = New-WordDocument -FilePath $FilePath
    }

    $ForestInformation = Get-WinADForest
    $ForestInformationTable = Get-WinADForestInformation -ForestInformation $ForestInformation


    $Toc = Add-WordToc -WordDocument $WordDocument -Title 'Table of content' -Switches C, A -RightTabPos 15 -HeaderStyle Heading1

    ### 1st section - Introduction
    $SectionScope = $WordDocument | Add-WordTocItem -Text 'Scope' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionScope = $WordDocument | Get-DocumentScope -Paragraph $SectionScope -CompanyName $CompanyName
    ### 2nd section - General Information
    # Leaving empty for now
    # $SectionGeneralInformation = $WordDocument | Add-WordTocItem -Text 'General Information' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1


    ### 3rd section - Forest Summary
    $SectionForestSummary = $WordDocument | Add-WordTocItem -Text 'General Information - Forest Summary' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionForestSummary = $WordDocument | Get-ForestSummary -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable
    #$SectionForestSummary = $WordDocument | Get-ForestFeatures -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ActiveDirectorySnapshot
    $SectionForestSummary = $WordDocument | Get-ForestFSMORoles -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable

    ### Section - UPN Summary
    # $SectionDomainUPNs = $WordDocument | Add-WordTocItem -Text 'General Information - Forest UPN Summary' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionForestSummary = $WordDocument | Add-WordParagraph
    $SectionForestSummary = $WordDocument | Get-ForestUPNSuffixes -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable -CompanyName $CompanyName
    $SectionForestSummary = $WordDocument | Add-WordParagraph
    $SectionForestSummary = $WordDocument | Get-ForestSPNSuffixes -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable -CompanyName $CompanyName

    foreach ($Domain in $ForestInformation.Domains) {
        $ADSnapshot = Get-ActiveDirectoryCleanData -Domain $Domain
        $ActiveDirectorySnapshot = Get-ActiveDirectoryProcessedData -ADSnapshot $ADSnapshot

        $SectionDomainSummary = $WordDocument | Add-WordTocItem -Text "General Information - Domain $Domain" -ListLevel 1 -ListItemType Numbered -HeadingType Heading1
        ### Section - Domain Summary
        $SectionDomainSummary = $WordDocument | Add-WordTocItem -Text "General Information - Domain Summary" -ListLevel 2 -ListItemType Numbered -HeadingType Heading2
        $SectionDomainSummary = $WordDocument | Get-DomainSummary -Paragraph $SectionDomainSummary -ActiveDirectorySnapshot $ActiveDirectorySnapshot -Domain $Domain
        ### Section - Password Policies
        $SectionPasswordPolicies = $WordDocument | Add-WordTocItem -Text 'General Information - Password Policies' -ListLevel 2 -ListItemType Numbered -HeadingType Heading2
        $SectionPasswordPolicies = $WordDocument | Get-DomainPasswordPolicies -Paragraph $SectionPasswordPolicies -ActiveDirectorySnapshot $ActiveDirectorySnapshot -Domain $Domain

        ### Section - Password Policies
        $SectionPasswordPolicies = $WordDocument | Add-WordTocItem -Text 'General Information - Group Policies' -ListLevel 2 -ListItemType Numbered -HeadingType Heading2
        $SectionPasswordPolicies = $WordDocument | Get-DomainGroupPolicies -Paragraph $SectionPasswordPolicies -ActiveDirectorySnapshot $ActiveDirectorySnapshot -Domain $Domain -Verbose
    }
    Save-WordDocument -WordDocument $WordDocument -Language 'en-US' -FilePath $FilePath -Supress $true
    if ($OpenDocument) { Invoke-Item $FilePath }
}

Clear-Host
Start-WinDocumentationServer -ComputerName 'AD1' -FilePathTemplate $FilePathTemplate -FilePath $FilePath -OpenDocument