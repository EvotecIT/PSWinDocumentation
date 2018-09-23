function New-DataBlock {
    param(
        [Xceed.Words.NET.Container]$WordDocument,
        [Object] $Section,
        [alias('Object')][Object] $Forest,
        [string] $Domain,
        [OfficeOpenXml.ExcelPackage] $Excel,
        [string] $SectionName,
        [nullable[bool]] $Sql # Global Sql Enable/Disable
    )
    if ($Section.Use) {
        if ($Domain) {
            $SectionDetails = "$Domain - $SectionName"
        } else {
            $SectionDetails = $SectionName
        }
        #Write-Verbose "New-ADDocumentBlock - Processing section [$Section][$($Section.SqlData)][Forest: $Forest][Domain: $Domain]"
        $TableData = Get-WinDocumentationData -DataToGet $Section.TableData -Object $Forest -Domain $Domain
        $ExcelData = Get-WinDocumentationData -DataToGet $Section.ExcelData -Object $Forest -Domain $Domain
        $ListData = Get-WinDocumentationData -DataToGet $Section.ListData -Object $Forest -Domain $Domain
        $SqlData = Get-WinDocumentationData -DataToGet $Section.SqlData -Object $Forest -Domain $Domain

        ### Preparing chart data
        $ChartData = (Get-WinDocumentationData -DataToGet $Section.ChartData -Object $Forest -Domain $Domain)
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

        #Write-Verbose "Table Data $($TableData.Count)"

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
                -TableColumnWidths $Section.TableColumnWidths `
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
                $ExcelWorksheet = Add-ExcelWorksheetData -ExcelDocument $Excel -ExcelWorksheetName $WorkSheetName -DataTable $ExcelData -AutoFit -AutoFilter -PreScanHeaders #-Verbose
                #| Convert-ToExcel -Path $Excel -AutoSize -AutoFilter -WorksheetName $WorkSheetName -ClearSheet -NoNumberConversion SSDL, GUID, ID, ACLs
            }
        }
        if ($Sql -and $Section.SQLExport -and $SqlData) {
            Write-Verbose "Sending [$SectionDetails] to SQL Server"
            $SqlQuery = Send-SqlInsert -Object $SqlData -SqlSettings $Section -Verbose
            foreach ($Query in $SqlQuery) {
                # if ($Query -like '*Error*') {
                Write-Color @script:WriteParameters -Text '[i] ', 'MS SQL Output: ', $Query -Color White, White, Yellow
                # }
            }
        }
    }
    if ($WordDocument) { return $WordDocument } else { return }
}