function Get-WinDocumentationData {
    param (
        [Object] $Data,
        [hashtable] $Forest,
        [string] $Domain
    )
    if ($Data -ne $null) {
        $Type = Get-ObjectType -Object $Data -ObjectName 'Get-WinDocumentationData' #-Verbose
        #Write-Verbose "Get-WinDocumentationData - Type: $($Type.ObjectTypeName) - Tabl $Data"
        if ($Type.ObjectTypeName -eq 'ActiveDirectory' -and $Data.ToString() -like 'Forest*') {
            return $Forest."$Data"
        } elseif ($Type.ObjectTypeName -eq 'ActiveDirectory' -and $Data.ToString() -like 'Domain*' ) {
            return $Forest.FoundDomains.$Domain."$Data"
        }
    }
    #Write-Verbose 'Get-WinDocumentationData - Data was $null'
    return
}
function Get-WinDocumentationText {
    param (
        [string[]] $Text,
        [hashtable] $Forest,
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
        [Xceed.Words.NET.Container]$WordDocument,
        [Object] $Section,
        [Object] $Forest,
        [string] $Domain,
        $Excel,
        [string] $SectionName
    )
    if ($Section.Use) {
        if ($Domain) {
            $SectionDetails = "$Domain - $SectionName"
        } else {
            $SectionDetails = $SectionName
        }
        #Write-Verbose "New-ADDocumentBlock - Processing section [$Section][$($Section.SqlData)][Forest: $Forest][Domain: $Domain]"
        $TableData = Get-WinDocumentationData -Data $Section.TableData -Forest $Forest -Domain $Domain
        $ExcelData = Get-WinDocumentationData -Data $Section.ExcelData -Forest $Forest -Domain $Domain
        $ListData = Get-WinDocumentationData -Data $Section.ListData -Forest $Forest -Domain $Domain
        $SqlData = Get-WinDocumentationData -Data $($Section.SqlData) -Forest $Forest -Domain $Domain

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
        $TextNoData = (Get-WinDocumentationText -Text $Section.TextNoData -Forest $Forest -Domain $Domain)

        if ($WordDocument) {
            Write-Verbose "Generating WORD Section for [$SectionDetails]"
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
                -TextNoData $TextNoData `
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
        }
        if ($Excel -and $Section.ExcelExport) {
            if ($Section.ExcelWorkSheet -eq '') {
                $WorkSheetName = $SectionDetails
            } else {
                $WorkSheetName = (Get-WinDocumentationText -Text $Section.ExcelWorkSheet -Forest $Forest -Domain $Domain)
            }
            if ($ExcelData) {
                Write-Verbose "Generating EXCEL Section for [$SectionDetails]"
                Add-ExcelWorksheetData -ExcelDocument $Excel -ExcelWorksheetName $WorkSheetName -DataTable $ExcelData -AutoFit -AutoFilter #-Verbose
                #| Convert-ToExcel -Path $Excel -AutoSize -AutoFilter -WorksheetName $WorkSheetName -ClearSheet -NoNumberConversion SSDL, GUID, ID, ACLs
            }
        }
        if ($Section.SQLExport -and $SqlData) {
            Write-Verbose "Sending [$SectionDetails] to SQL Server"
            $SqlQuery = Send-SqlInsert -Object $SqlData -SqlSettings $Section
            foreach ($Query in $SqlQuery) {
                Write-Color @script:WriteParameters -Text '[i] ', 'MS SQL Output: ', $Query -Color White, White, Yellow
            }
        }
    }
    if ($WordDocument) { return $WordDocument } else { return }
}

function Start-ActiveDirectoryDocumentation {
    [CmdletBinding()]
    param (
        [string] $FilePath,
        [string] $FilePathExcel,
        [switch] $CleanDocument,
        [string] $CompanyName = 'Evotec',
        [switch] $OpenDocument,
        [switch] $OpenExcel
    )
    #Write-Warning 'This is legacy command. Use Start-Documentation instead.'
    # Left here for legacy reasons.
    $Document = $Script:Document
    $Document.Configuration.Prettify.CompanyName = $CompanyName
    if ($CleanDocument) {
        $Document.Configuration.Prettify.UseBuiltinTemplate = $false
    }
    $Document.Configuration.Options.OpenDocument = $OpenDocument
    $Document.Configuration.Options.OpenExcel = $OpenExcel
    $Document.DocumentAD.FilePathWord = $FilePath
    if ($FilePathExcel) {
        $Document.DocumentAD.ExportExcel = $true
        $Document.DocumentAD.FilePathExcel = $FilePathExcel
    } else {
        $Document.DocumentAD.ExportExcel = $false
    }

    Start-Documentation -Document $Document
}