function Get-WinDocumentationData {
    param (
        [Object] $Data,
        [Object] $Forest,
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
            GroupPoliciesDetails { return $Forest.FoundDomains.$Domain.GroupPoliciesDetails }
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
        [string[]] $Text,
        [Object] $Forest,
        [string] $Domain
    )
    $Array = @()
    foreach ($T in $Text) {
        $T = $T.Replace('<CompanyName>', $Document.Configuration.Prettify.CompanyName)
        $T = $T.Replace('<ForestName>', $Forest.ForestName)
        $T = $T.Replace('<ForestNameDN>', $Forest.RootDSE.defaultNamingContext)
        $T = $T.Replace('<Domain>', $Domain)
        $T = $T.Replace('<DomainNetBios>', $Forest.FoundDomains.$Domain.DomainInformation.NetBIOSName)
        $T = $T.Replace('<DomainDN>', $Forest.FoundDomains.$Domain.DomainInformation.DistinguishedName)
        $Array += $T
    }
    return $Array
}

function New-ADDocumentBlock {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Words.NET.Container]$WordDocument,
        [Object] $Section,
        [Object] $Forest,
        [string] $Domain,
        [string] $Excel
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

        $ListBuilderContent = (Get-WinDocumentationText -Text $Section.ListBuilderContent -Forest $Forest -Domain $Domain)

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
            -TableMaximumColumns $Section.TableMaximumColumns `
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
            -ChartValues $ChartValues `
            -ListBuilderContent $ListBuilderContent `
            -ListBuilderType $Section.ListBuilderType `
            -ListBuilderLevel $Section.ListBuilderLevel

        if ($TableData) {
            #$WorkSheetName = [System.IO.Path]::GetRandomFileName()
            #$TableData | Convert-ToExcel -Path $Excel -AutoSize -AutoFilter -WorksheetName $WorkSheetName -ClearSheet -NoNumberConversion * #SSDL, GUID, ID, ACLs
        }
    }
    return $WordDocument
}

function Start-ActiveDirectoryDocumentation {
    [CmdletBinding()]
    param (
        [string] $FilePath,
        [switch] $CleanDocument,
        [string] $CompanyName = 'Evotec',
        [switch] $OpenDocument
    )
    # Left here for legacy reasons.
    $Document = $Script:Document
    $Document.Configuration.Prettify.CompanyName = $CompanyName
    if ($CleanDocument) {
        $Document.Configuration.Prettify.UseBuiltinTemplate = $false
    }
    $Document.Configuration.Options.OpenDocument = $OpenDocument
    $Document.DocumentAD.FilePathWord = $FilePath

    Start-Documentation -Document $Document
}