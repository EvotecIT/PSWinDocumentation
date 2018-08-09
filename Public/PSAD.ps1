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
        $Data,
        $Forest,
        [string] $Domain
    )
    $Type = Get-ObjectType $Data
    #Write-Verbose "Get-WinDocumentationData - Type: $($Type.ObjectTypeName) - Tabl"
    if ($Type.ObjectTypeName -eq 'Forest') {
        switch ( $Data ) {
            Summary { return $Forest.ForestInformation }
            FSMO { return $Forest.FSMO }
            OptionalFeatures { return $Forest.OptionalFeatures }
            UPNSuffixes { return $Forest.UPNSuffixes }
            SPNSuffixes {
                write-verbose 'spn suffixes'
                return $Forest.SPNSuffixes
            }
            Sites { return $Forest.Sites }
            Sites1 { return $Forest.Sites1 }
            Sites2 { return $Forest.Sites2 }
            Subnets { return $Forest.Subnets }
            Subnets1 { return $Forest.Subnets1 }
            Subnets2 { return $Forest.Subnets2 }
            SiteLinks { return $Forest.SiteLinks }
            default { return $null }
        }
    } elseif ($Type.ObjectTypeName -eq 'Domain' ) {
        switch ( $Data ) {
            DomainControllers { return $Forest.FoundDomains.$Domain.DomainControllers }
        }
    }
}
function Get-WinDocumentationText {
    param (
        [string] $Text,
        $Forest
    )
    #$ForestInformation.GetType()
    $Text = $Text.Replace('<CompanyName>', $Document.Configuration.Prettify.CompanyName)
    $Text = $Text.Replace('<ForestName>', $Forest.ForestName)
    return $Text
}

function New-ADDocumentBlock {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Words.NET.Container]$WordDocument,
        $Section,
        $Forest,
        [string] $Domain
    )
    if ($Section.Use) {
        #Write-Verbose "New-ADDocumentBlock - Processing section [$Section][$($Section.TableData)]"
        $WordDocument | New-WordBlock `
            -TocGlobalDefinition $Section.TocGlobalDefinition`
            -TocGlobalTitle $Section.TocGlobalTitle `
            -TocGlobalSwitches $Section.TocGlobalSwitches `
            -TocGlobalRightTabPos $Section.TocGlobalRightTabPos `
            -TocEnable $Section.TocEnable `
            -TocText (Get-WinDocumentationText -Text $Section.TocText -Forest $Forest) `
            -TocListLevel $Section.TocListLevel `
            -TocListItemType $Section.TocListItemType `
            -TocHeadingType $Section.TocHeadingType `
            -TableData (Get-WinDocumentationData -Data $Section.TableData -Forest $Forest -Domain $Domain) `
            -TableDesign $Section.TableDesign `
            -TableTitleMerge $Section.TableTitleMerge `
            -TableTitleText (Get-WinDocumentationText -Text $Section.TableTitleText -Forest $Forest) `
            -Text (Get-WinDocumentationText -Text $Section.Text -Forest $Forest) `
            -EmptyParagraphsBefore $Section.EmptyParagraphsBefore `
            -EmptyParagraphsAfter $Section.EmptyParagraphsAfter `
            -PageBreaksBefore $Section.PageBreaksBefore `
            -PageBreaksAfter $Section.PageBreaksAfter `
            -TextAlignment $Section.TextAlignment `
            -ListData (Get-WinDocumentationData -Data $Section.ListData -Forest $Forest -Domain $Domain) `
            -ListType $Section.ListType `
            -ListTextEmpty $Section.ListTextEmpty #-Verbose

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
        $Forest = Get-WinADForestInformation

        ### Starting WORD
        $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentAD.FilePathWord

        ### Start Sections
        $ADSectionsForest = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionForest
        foreach ($Section in $ADSectionsForest) {
            Write-Verbose "Generating WORD Section for [$Section]"
            $WordDocument = $WordDocument | New-ADDocumentBlock -Section $Document.DocumentAD.Sections.SectionForest.$Section -Forest $Forest
        }
        foreach ($Domain in $Forest.Domains) {
            $ADSectionsDomain = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionDomain
            foreach ($Section in $ADSectionsDomain) {
                Write-Verbose "Generating WORD Section for [$Domain - $Section]"
                $WordDocument = $WordDocument | New-ADDocumentBlock -Section $Document.DocumentAD.Sections.SectionDomain.$Section -Forest $Forest -Domain $Domain
            }
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