function Start-Documentation {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    Test-ModuleAvailability
    Test-ForestConnectivity
    Test-Configuration -Document $Document

    if ($Document.DocumentAD.Enable) {
        $TypesRequired = Get-TypesRequired -Document $Document
        $ADSectionsForest = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionForest
        $ADSectionsDomain = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionDomain
        $Forest = Get-WinADForestInformation -TypesRequired $TypesRequired


        ### Starting WORD
        $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentAD.FilePathWord
        $ExcelDocument = $Document.DocumentAD.FilePathExcel

        ### Start Sections

        foreach ($Section in $ADSectionsForest) {
            Write-Verbose "Generating WORD Section for [$Section]"
            $WordDocument = New-ADDocumentBlock `
                -WordDocument $WordDocument `
                -Section $Document.DocumentAD.Sections.SectionForest.$Section `
                -Forest $Forest `
                -Excel $ExcelDocument
            #$ExcelDocument = $ExcelDocument | New-ExportExcelBlock -Section $Document.DocumentAD.Sections.SectionDomain.$Section -Forest $Forest -Domain $Domain
        }
        foreach ($Domain in $Forest.Domains) {
            foreach ($Section in $ADSectionsDomain) {
                Write-Verbose "Generating WORD Section for [$Domain - $Section]"
                $WordDocument = New-ADDocumentBlock `
                    -WordDocument $WordDocument `
                    -Section $Document.DocumentAD.Sections.SectionDomain.$Section `
                    -Forest $Forest `
                    -Domain $Domain `
                    -Excel $ExcelDocument
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
    }
}
