Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord #-Force
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

    $WordDocument | Add-WordSection -PageBreak

    ### 1st section - Introduction
    $SectionScope = $WordDocument | Add-WordTocItem -Text 'Scope' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionScope = $WordDocument | Get-DocumentScope -Paragraph $SectionScope -CompanyName $CompanyName

    $WordDocument | Add-WordSection -PageBreak
    ### 3rd section - Forest Summary
    $SectionForestSummary = $WordDocument | Add-WordTocItem -Text 'General Information - Forest Summary' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $SectionForestSummary = $WordDocument | Get-ForestSummary -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable
    $SectionForestSummary = $WordDocument | Get-ForestFSMORoles -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable

    $WordDocument | New-WordBuildingBlock `
        -TableData $ForestInformationTable.OptionalFeatures `
        -TableDesign ColorfulGridAccent5 `
        -TableTitleMerge $true `
        -TableTitleText 'Optional Features' `
        -Text "Following table contains optional forest features"


    ### Section - UPN Summary
    $SectionForestSummary = $WordDocument | Add-WordParagraph
    $SectionForestSummary = $WordDocument | Get-ForestUPNSuffixes -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable -CompanyName $CompanyName
    $SectionForestSummary = $WordDocument | Add-WordParagraph
    $SectionForestSummary = $WordDocument | Get-ForestSPNSuffixes -Paragraph $SectionForestSummary -ActiveDirectorySnapshot $ForestInformationTable -CompanyName $CompanyName

    foreach ($Domain in $ForestInformation.Domains) {
        $WordDocument | Add-WordSection -PageBreak
        $ADSnapshot = Get-ActiveDirectoryCleanData -Domain $Domain
        $ActiveDirectorySnapshot = Get-ActiveDirectoryProcessedData -ADSnapshot $ADSnapshot

        $SectionDomainSummary = $WordDocument | Add-WordTocItem -Text "General Information - Domain $Domain" -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
        ### Section - Domain Summary
        $SectionDomainSummary = $WordDocument | Add-WordTocItem -Text "General Information - Domain Summary" -ListLevel 1 -ListItemType Numbered -HeadingType Heading2
        $SectionDomainSummary = $WordDocument | Get-DomainSummary -Paragraph $SectionDomainSummary -ActiveDirectorySnapshot $ActiveDirectorySnapshot -Domain $Domain

        $WordDocument | New-WordBuildingBlock `
            -TocEnable $True `
            -TocText 'General Information - Password Policies' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $ActiveDirectorySnapshot.DefaultPassWordPoLicy `
            -TableDesign ColorfulGridAccent5 `
            -TableTitleMerge $True `
            -TableTitleText "Default Password Policy for $Domain" `
            -Text 'Following table contains password policies'

        $WordDocument | New-WordBuildingBlock `
            -TocEnable $True `
            -TocText 'General Information - Group Policies' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $ActiveDirectorySnapshot.GroupPoliciesTable `
            -TableDesign ColorfulGridAccent5 `
            -Text "Following table contains group policies for $Domain"

        $WordDocument | New-WordBuildingBlock `
            -TocEnable $True `
            -TocText 'General Information - Priviliged Members' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $ActiveDirectorySnapshot.PriviligedGroupMembers `
            -TableDesign ColorfulGridAccent5

    }
    Save-WordDocument -WordDocument $WordDocument -Language 'en-US' -FilePath $FilePath -Supress $true
    if ($OpenDocument) { Invoke-Item $FilePath }
}

Clear-Host
Start-WinDocumentationServer -ComputerName 'AD1' -FilePathTemplate $FilePathTemplate -FilePath $FilePath -OpenDocument