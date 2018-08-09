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
            SPNSuffixes { return $Forest.SPNSuffixes }
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
            DomainInformation { return $Forest.FoundDomains.$Domain.DomainInformation }
            FSMO { return $Forest.FoundDomains.$Domain.FSMO }
            DefaultPasswordPoLicy { return $Forest.FoundDomains.$Domain.DefaultPasswordPoLicy }
            GroupPolicies { return $Forest.FoundDomains.$Domain.GroupPolicies }
            OrganizationalUnits { return $Forest.FoundDomains.$Domain.OrganizationalUnits }
            PriviligedGroupMembers { return $Forest.FoundDomains.$Domain.PriviligedGroupMembers }
            DomainAdministrators { return $Forest.FoundDomains.$Domain.DomainAdministrators }
            Users { return $Forest.FoundDomains.$Domain.Users }
            UsersCount { return $Forest.FoundDomains.$Domain.UsersCount }
        }
    }
}
function Get-WinDocumentationText {
    param (
        [string] $Text,
        $Forest,
        [string] $Domain
    )
    #$ForestInformation.GetType()
    $Text = $Text.Replace('<CompanyName>', $Document.Configuration.Prettify.CompanyName)
    $Text = $Text.Replace('<ForestName>', $Forest.ForestName)
    $Text = $Text.Replace('<Domain>', $Domain)
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
        $TableData = (Get-WinDocumentationData -Data $Section.TableData -Forest $Forest -Domain $Domain)
        $ListData = (Get-WinDocumentationData -Data $Section.ListData -Forest $Forest -Domain $Domain)

        ### Preparing chart data
        $ChartData = (Get-WinDocumentationData -Data $Section.ChartData -Forest $Forest -Domain $Domain)
        if ($ChartData) {
            if ($Section.ChartKeys -eq 'Keys' -and $Section.ChartValues -eq 'Values') {
                $ChartKeys = (Convert-KeyToKeyValue $ChartData).Keys
                $ChartValues = (Convert-KeyToKeyValue $ChartData).Values
            } else {
                $ChartKeys = (Convert-TwoArraysIntoOne -Object $ChartData.($Section.ChartKeys[0]) -ObjectToAdd $ChartData.($Section.ChartKeys[1]))
                $ChartValues = ($ChartData.($Section.ChartValues))
            }
        }

        ### Converts for Text
        $TocText = (Get-WinDocumentationText -Text $Section.TocText -Forest $Forest -Domain $Domain)
        $TableTitleText = (Get-WinDocumentationText -Text $Section.TableTitleText -Forest $Forest -Domain $Domain)
        $Text = (Get-WinDocumentationText -Text $Section.Text -Forest $Forest -Domain $Domain)
        $ChartTitle = (Get-WinDocumentationText -Text $Section.ChartTitle -Forest $Forest -Domain $Domain)

        $WordDocument | New-WordBlock `
            -TocGlobalDefinition $Section.TocGlobalDefinition`
            -TocGlobalTitle $Section.TocGlobalTitle `
            -TocGlobalSwitches $Section.TocGlobalSwitches `
            -TocGlobalRightTabPos $Section.TocGlobalRightTabPos `
            -TocEnable $Section.TocEnable `
            -TocText $TocText `
            -TocListLevel $Section.TocListLevel `
            -TocListItemType $Section.TocListItemType `
            -TocHeadingType $Section.TocHeadingType `
            -TableData $TableData `
            -TableDesign $Section.TableDesign `
            -TableTitleMerge $Section.TableTitleMerge `
            -TableTitleText $TableTitleText `
            -Text $Text `
            -EmptyParagraphsBefore $Section.EmptyParagraphsBefore `
            -EmptyParagraphsAfter $Section.EmptyParagraphsAfter `
            -PageBreaksBefore $Section.PageBreaksBefore `
            -PageBreaksAfter $Section.PageBreaksAfter `
            -TextAlignment $Section.TextAlignment `
            -ListData $ListData `
            -ListType $Section.ListType `
            -ListTextEmpty $Section.ListTextEmpty `
            -ChartEnable $Section.ChartEnable `
            -ChartTitle $ChartTitle `
            -ChartKeys $ChartKeys `
            -ChartValues $ChartValues
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
}