function Register-DataFillers {
    param(
        $Document,
        $ForestInformation
    )

    $ForestName = $ForestInformation.ForestName
    $CompanyName = $Document.Configuration.CompanyName
}

function Test-File {
    param(
        [string] $File,
        [string] $FileName,
        [switch] $Require,
        [switch] $Skip
    )
    [int] $ErrorCount = 0
    if ($Skip) {
        return $ErrorCount
    }
    if ($File -ne '') {
        if ($Require) {
            if (Test-Path $File) {
                return $ErrorCount
            } else {
                Write-Color  @Script:WriteParameters '[e] ', $FileName, " doesn't exists (", $File, "). It's required if you want to use this feature." -Color Red, Yellow, Yellow, White
                $ErrorCount++
            }
        }
    } else {
        $ErrorCount++
        Write-Color @Script:WriteParameters '[e] ', $FileName, " was empty. It's required if you want to use this feature." -Color Red, Yellow, White
    }
    return $ErrorCount
}

function Test-Configuration {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    [int] $ErrorCount = 0
    $Script:WriteParameters = $Document.Configuration.DisplayConsole


    $Keys = Get-ObjectKeys -Object $Document -Ignore 'Configuration'
    foreach ($Key in $Keys) {
        $ErrorCount += Test-File -File $Document.$Key.FilePathWord -FileName 'FilePathWord' -Skip:(-not $Document.$Key.ExportWord)
        $ErrorCount += Test-File -File $Document.$Key.FilePathExcel -FileName 'FilePathExcel' -Skip:(-not $Document.$Key.ExportExcel)
    }
    if ($ErrorCount -ne 0) {
        Exit
    }
}
function Get-DocumentPath {
    [CmdletBinding()]
    param (
        [System.Object] $Document,
        [string] $FinalDocumentLocation
    )
    if ($Document.Configuration.Prettify.UseBuiltinTemplate) {
        Write-Verbose 'Get-DocumentPath - Option 1'
        $WordDocument = Get-WordDocument -FilePath "$((get-item $PSScriptRoot).Parent.FullName)\Templates\WordTemplate.docx"
    } else {
        if ($Document.Configuration.Prettify.CustomTemplatePath) {
            if (Test-File -File $Document.Configuration.Prettify.CustomTemplatePath -FileName 'CustomTemplatePath' -eq 0) {
                Write-Verbose 'Get-DocumentPath - Option 2'
                $WordDocument = Get-WordDocument -FilePath $Document.Configuration.Prettify.CustomTemplatePath
            } else {
                Write-Verbose 'Get-DocumentPath - Option 3'
                $WordDocument = New-WordDocument -FilePath $FinalDocumentLocation
            }
        } else {
            Write-Verbose 'Get-DocumentPath - Option 4'
            $WordDocument = New-WordDocument -FilePath $FinalDocumentLocation
        }
    }
    if ($WordDocument -eq $null) { Write-Verbose ' Null'}
    return $WordDocument
}
function Get-WinDocumentationData {
    param (
        [nullable[TableData]] $TableData,
        $ForestInformation
    )
    switch ( $TableData ) {
        ForestSummary { return $ForestInformation.ForestInformation }
        ForestFSMO { return $ForestInformation.FSMO }
        ForestOptionalFeatures { return $ForestInformation.OptionalFeatures }
        ForestUPNSuffixes { return $ForestInformation.UPNSuffixes }
        ForestSPNSuffixes {
            write-verbose 'spn suffixes'
            return $ForestInformation.SPNSuffixes
        }
        ForestSites { return $ForestInformation.Sites }
        ForestSites1 { return $ForestInformation.Sites1 }
        ForestSites2 { return $ForestInformation.Sites2 }
        ForestSubnets { return $ForestInformation.Subnets }
        ForestSubnets1 { return $ForestInformation.Subnets1 }
        ForestSubnets2 { return $ForestInformation.Subnets2 }
        ForestSiteLinks { return $ForestInformation.SiteLinks }
        default { return $null }
    }
}
function Get-WinDocumentationText {
    param (
        [string] $Text,
        $ForestInformation
    )
    #$ForestInformation.GetType()
    $Text = $Text.Replace('<CompanyName>', $Document.Configuration.Prettify.CompanyName)
    $Text = $Text.Replace('<ForestName>', $ForestInformation.ForestName)
    return $Text
}

function New-ADDocumentBlock {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Words.NET.Container]$WordDocument,
        $Section,
        $ForestInformation
    )
    if ($Section.Use) {
        $WordDocument | New-WordBlock `
            -TocGlobalDefinition $Section.TocGlobalDefinition`
            -TocGlobalTitle $Section.TocGlobalTitle `
            -TocGlobalSwitches $Section.TocGlobalSwitches `
            -TocGlobalRightTabPos $Section.TocGlobalRightTabPos `
            -TocEnable $Section.TocEnable `
            -TocText (Get-WinDocumentationText -Text $Section.TocText -ForestInformation $ForestInformation) `
            -TocListLevel $Section.TocListLevel `
            -TocListItemType $Section.TocListItemType `
            -TocHeadingType $Section.TocHeadingType `
            -TableData (Get-WinDocumentationData -TableData $Section.TableData -ForestInformation $ForestInformation) `
            -TableDesign $Section.TableDesign `
            -TableTitleMerge $Section.TableTitleMerge `
            -TableTitleText (Get-WinDocumentationText -Text $Section.TableTitleText -ForestInformation $ForestInformation) `
            -Text (Get-WinDocumentationText -Text $Section.Text -ForestInformation $ForestInformation) `
            -EmptyParagraphsBefore $Section.EmptyParagraphsBefore `
            -EmptyParagraphsAfter $Section.EmptyParagraphsAfter `
            -PageBreaksBefore $Section.PageBreaksBefore `
            -PageBreaksAfter $Section.PageBreaksAfter `
            -TextAlignment $Section.TextAlignment `
            -ListData (Get-WinDocumentationData -TableData $Section.ListData -ForestInformation $ForestInformation) `
            -ListType $Section.ListType `
            -ListTextEmpty $Section.ListTextEmpty `
            -Verbose

    }
    return $WordDocument
}

function Search-Command($CommandName) {
    return [bool](Get-Command -Name $CommandName -ErrorAction SilentlyContinue)
}

function Test-ModuleAvailability {
    if (Search-Command -CommandName 'Get-AdForest') {
        # future use
    } else {
        Write-Warning 'Modules required to run not found.'
        Exit
    }
}

function Start-Documentation {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    Test-ModuleAvailability
    Test-Configuration -Document $Document



    if ($Document.DocumentAD.Enable) {
        $ForestInformation = Get-WinADForestInformation
        #$ForestInformation.FoundDomains.Count

        ### Starting WORD
        $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentAD.FilePathWord

        ### Start Sections
        $ADSectionsForest = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionForest
        foreach ($Section in $ADSectionsForest) {
            Write-Verbose "Generating WORD Section for [$Section]"
            $WordDocument = $WordDocument | New-ADDocumentBlock -Section $Document.DocumentAD.Sections.SectionForest.$Section -ForestInformation $ForestInformation
        }
        ### End Sections

        ### Ending WORD
        $FilePath = Save-WordDocument -WordDocument $WordDocument `
            -Language $Document.Configuration.Prettify.Language `
            -FilePath $Document.DocumentAD.FilePathWord `
            -Supress $True `
            -OpenDocument:$Document.Configuration.Options.OpenDocument
    }
    return

    Write-Verbose 'Start-ActiveDirectoryDocumentation - Working...2'
    foreach ($Domain in $ForestInformation.Domains) {
        $WordDocument | Add-WordPageBreak -Supress $True
        Write-Verbose 'Start-ActiveDirectoryDocumentation - Getting domain information'
        $DomainInformation = Get-WinADDomainInformation -Domain $Domain

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
            -TableTitleMerge $False `
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
            $DomainInformation = Get-WinADDomainInformation -Domain $Domain
            $DomainInformation.DomainControllers  | Export-Excel -AutoSize -Path $FilePathExcel -AutoFilter -WorkSheetname "$Domain DCs" -ClearSheet -FreezeTopRow
            $DomainInformation.GroupPoliciesDetails | Export-Excel -AutoSize -Path $FilePathExcel -AutoFilter -WorksheetName "$Domain GPOs Details" -ClearSheet -FreezeTopRow -NoNumberConversion SSDL, GUID, ID, ACLs
            $DomainInformation.GroupPoliciesDetails | fl *
        }

    }
    if ($OpenDocument) { Invoke-Item $FilePath }
    if ($OpenWorkbook) { Invoke-Item $FilePathExcel }
}