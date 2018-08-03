function Start-ActiveDirectoryDocumentation {
    [CmdletBinding()]
    param (
        [string] $FilePath,
        [switch] $OpenDocument,
        [switch] $CleanDocument,
        [string] $CompanyName = 'Evotec',
        [string] $FilePathExcel
    )
    if ($FilePath -eq '') { throw 'FilePath is required. This should be path where you want to save your document to.'}

    $FilePathTemplate = "$((get-item $PSScriptRoot).Parent.FullName)\Templates\WordTemplate.docx"

    if ($CleanDocument) {
        $WordDocument = New-WordDocument -FilePath $FilePath
    } else {
        $WordDocument = Get-WordDocument -FilePath $FilePathTemplate
    }

    Write-Verbose 'Start-ActiveDirectoryDocumentation - Getting Forest Information'
    $ForestInformation = Get-WinADForestInformation
    Write-Verbose 'Start-ActiveDirectoryDocumentation - Working...1'
    $Toc = Add-WordToc -WordDocument $WordDocument -Title 'Table of content' -Switches C, A -RightTabPos 15

    $WordDocument | Add-WordPageBreak -Supress $True

    ### 1st section - Introduction
    $Text = "This document provides a low-level design of roles and permissions for the IT infrastructure team at $CompanyName organization. This document utilizes knowledge from AD General Concept document that should be delivered with this document. Having all the information described in attached document one can start designing Active Directory with those principles in mind. It's important to know while best practices that were described are important in decision making they should not be treated as final and only solution. Most important aspect is to make sure company has full usability of Active Directory and is happy with how it works. Making things harder just for the sake of implementation of best practices isn't always the best way to go."
    $WordDocument | New-WordBlock `
        -TocEnable $True `
        -TocText 'Scope' `
        -TocListLevel 0 `
        -TocListItemType Numbered `
        -TocHeadingType Heading1 `
        -Text $Text

    $WordDocument | Add-WordPageBreak -Supress $True

    ### Section - Forest Summary
    $WordDocument | New-WordBlockTable `
        -TocEnable $True `
        -TocText 'General Information - Forest Summary' `
        -TocListLevel 0 `
        -TocListItemType Numbered `
        -TocHeadingType Heading1 `
        -TableData $ForestInformation.ForestInformation `
        -TableDesign ColorfulGridAccent5 `
        -TableTitleMerge $True `
        -TableTitleText "Forest Summary" `
        -Text  "Active Directory at $CompanyName has a forest name $($ForestInformation.ForestName). Following table contains forest summary with important information:"

    $WordDocument | New-WordBlockTable `
        -TableData $ForestInformation.FSMO `
        -TableDesign ColorfulGridAccent5 `
        -TableTitleMerge $true `
        -TableTitleText 'FSMO Roles' `
        -Text 'Following table contains FSMO servers' `
        -EmptyParagraphsBefore 1

    $WordDocument | New-WordBlockTable `
        -TableData $ForestInformation.OptionalFeatures `
        -TableDesign ColorfulGridAccent5 `
        -TableTitleMerge $true `
        -TableTitleText 'Optional Features' `
        -Text "Following table contains optional forest features" `
        -EmptyParagraphsBefore 1

    ### Section - UPN Summary
    $WordDocument | New-WordBlockList `
        -Text "Following UPN suffixes were created in this forest:" `
        -TextListEmpty "No UPN suffixes were created in this forest." `
        -ListType Bulleted `
        -ListData $ForestInformation.UPNSuffixes `
        -EmptyParagraphsBefore 1

    Write-Verbose 'Start-ActiveDirectoryDocumentation - Section Forest UPN'
    $WordDocument | New-WordBlockList `
        -Text "Following SPN suffixes were created in this forest:" `
        -TextListEmpty "No SPN suffixes were created in this forest." `
        -ListType Bulleted `
        -ListData $ForestInformation.SPNSuffixes `
        -EmptyParagraphsBefore 1

    Write-Verbose 'Start-ActiveDirectoryDocumentation - Section Forest Sites'
    $WordDocument | New-WordBlockTable `
        -TocEnable $True `
        -TocText 'General Information - Sites' `
        -TocListLevel 1 `
        -TocListItemType Numbered `
        -TocHeadingType Heading1 `
        -TableData $ForestInformation.Sites1 `
        -TableDesign ColorfulGridAccent5 `
        -Text  "Forest Sites list can be found below"

    $WordDocument | New-WordBlockTable `
        -TocHeadingType Heading1 `
        -TableData $ForestInformation.Sites2 `
        -TableDesign ColorfulGridAccent5 `
        -Text  "Forest Sites list can be found below" `
        -EmptyParagraphsBefore 1


    Write-Verbose 'Start-ActiveDirectoryDocumentation - Section Forest Subnets'
    $WordDocument | New-WordBlockTable `
        -TocEnable $True `
        -TocText 'General Information - Subnets' `
        -TocListLevel 1 `
        -TocListItemType Numbered `
        -TocHeadingType Heading1 `
        -TableData $ForestInformation.Subnets1 `
        -TableDesign ColorfulGridAccent5 `
        -Text  "Forest Subnet information is available below"


    $WordDocument | New-WordBlockTable `
        -TocHeadingType Heading1 `
        -TableData $ForestInformation.Subnets2 `
        -TableDesign ColorfulGridAccent5 `
        -Text  "Table below contains information regarding relation between Subnets and sites" `
        -EmptyParagraphsBefore 1

    Write-Verbose 'Start-ActiveDirectoryDocumentation - Section Forest SiteLinks'
    $WordDocument | New-WordBlockTable `
        -TocEnable $True `
        -TocText 'General Information - Site Links' `
        -TocListLevel 1 `
        -TocListItemType Numbered `
        -TocHeadingType Heading1 `
        -TableData $ForestInformation.SiteLinks `
        -TableDesign ColorfulGridAccent5 `
        -Text  "Forest Site Links information is available in table below"


    Write-Verbose 'Start-ActiveDirectoryDocumentation - Working...2'
    foreach ($Domain in $ForestInformation.Domains) {
        $WordDocument | Add-WordPageBreak -Supress $True
        Write-Verbose 'Start-ActiveDirectoryDocumentation - Getting domain information'
        $DomainInformation = Get-WinDomainInformation -Domain $Domain

        $SectionDomainSummary = $WordDocument | Add-WordTocItem -Text "General Information - Domain $Domain" -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
        ### Section - Domain Summary
        $SectionDomainSummary = $WordDocument | Add-WordTocItem -Text "General Information - Domain Summary" -ListLevel 1 -ListItemType Numbered -HeadingType Heading2
        $SectionDomainSummary = $WordDocument | Get-DomainSummary -Paragraph $SectionDomainSummary -ActiveDirectorySnapshot $DomainInformation.ADSnapshot -Domain $Domain

        Write-Verbose "Start-ActiveDirectoryDocumentation - Creating section for $Domain - Domain Controllers"

        $WordDocument | New-WordBlockTable `
            -TocEnable $True `
            -TocText 'General Information - Domain Controllers' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $DomainInformation.DomainControllers `
            -TableDesign ColorfulGridAccent5 `
            -TableMaximumColumns 8 `
            -Text 'Following table contains domain controllers'

        Write-Verbose "Start-ActiveDirectoryDocumentation - Creating section for $Domain - FSMO Roles"

        $WordDocument | New-WordBlockTable `
            -TableData $DomainInformation.FSMO `
            -TableDesign ColorfulGridAccent5 `
            -TableTitleMerge $true `
            -TableTitleText "FSMO Roles for $Domain" `
            -Text "Following table contains FSMO servers with roles for domain $Domain" `
            -EmptyParagraphsBefore 1

        $WordDocument | New-WordBlockTable `
            -TocEnable $True `
            -TocText 'General Information - Password Policies' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $DomainInformation.DefaultPassWordPoLicy `
            -TableDesign ColorfulGridAccent5 `
            -TableTitleMerge $True `
            -TableTitleText "Default Password Policy for $Domain" `
            -Text 'Following table contains password policies'

        $WordDocument | New-WordBlockTable `
            -TocEnable $True `
            -TocText 'General Information - Group Policies' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $DomainInformation.GroupPolicies `
            -TableDesign ColorfulGridAccent5 `
            -Text "Following table contains group policies for $Domain"

        $WordDocument | New-WordBlockTable `
            -TocEnable $True `
            -TocText 'General Information - Organizational Units' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $DomainInformation.OrganizationalUnits `
            -TableDesign ColorfulGridAccent5 `
            -Text "Following table contains all OU's created in $Domain"

        $WordDocument | New-WordBlockTable `
            -TocEnable $True `
            -TocText 'General Information - Priviliged Members' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $DomainInformation.PriviligedGroupMembers `
            -TableDesign ColorfulGridAccent5 `
            -Text 'Following table contains list of priviliged groups and count of the members in it.' `
            -ChartEnable $True `
            -ChartTitle 'Priviliged Group Members' `
            -ChartKeys (Convert-TwoArraysIntoOne -Object $DomainInformation.PriviligedGroupMembers.'Group Name' -ObjectToAdd $DomainInformation.PriviligedGroupMembers.'Members Count') `
            -ChartValues ($DomainInformation.PriviligedGroupMembers.'Members Count')

        $WordDocument | New-WordBlockTable `
            -TocEnable $True `
            -TocText 'General Information - Domain Administrators' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $DomainInformation.DomainAdministrators `
            -TableDesign ColorfulGridAccent5 `
            -Text 'Following users have highest domain priviliges and are able to control a lot of Windows resources.'

        $WordDocument | New-WordBlockTable `
            -TocEnable $True `
            -TocText 'General Information - Users Count' `
            -TocListLevel 1 `
            -TocListItemType Numbered `
            -TocHeadingType Heading2 `
            -TableData $DomainInformation.UsersCount `
            -TableDesign ColorfulGridAccent5 `
            -TableTitleMerge $True `
            -TableTitleText 'Users Count' `
            -Text "Following table and chart shows number of users in its categories" `
            -ChartEnable $True `
            -ChartTitle 'Users Count' `
            -ChartKeys (Convert-KeyToKeyValue $DomainInformation.UsersCount).Keys `
            -ChartValues (Convert-KeyToKeyValue $DomainInformation.UsersCount).Values

    }

    Save-WordDocument -WordDocument $WordDocument -Language 'en-US' -FilePath $FilePath -Supress $true #-Verbose

    if ($FilePathExcel) {
        $ForestInformation.ForestInformation | Export-Excel -AutoSize -Path $FilePathExcel -AutoFilter -Verbose -WorkSheetname 'Forest Information' -ClearSheet -FreezeTopRow
        $ForestInformation.FSMO | Export-Excel -AutoSize -Path $FilePathExcel -AutoFilter -WorkSheetname 'Forest FSMO' -FreezeTopRow
        foreach ($Domain in $ForestInformation.Domains) {
            $DomainInformation = Get-WinDomainInformation -Domain $Domain
            $DomainInformation.DomainControllers  | Export-Excel -AutoSize -Path $FilePathExcel -AutoFilter -WorkSheetname "$Domain DCs" -ClearSheet -FreezeTopRow
        }

    }
    if ($OpenDocument) { Invoke-Item $FilePath }
}