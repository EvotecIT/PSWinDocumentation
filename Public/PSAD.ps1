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
        if ($Document.DocumentAD.ExportExcel) {
            $ExcelDocument = New-ExcelDocument
        }

        ### Start Sections
        foreach ($Section in $ADSectionsForest) {
            $WordDocument = New-ADDocumentBlock `
                -WordDocument $WordDocument `
                -Section $Document.DocumentAD.Sections.SectionForest.$Section `
                -Forest $Forest `
                -Excel $ExcelDocument `
                -SectionName $Section
        }
        foreach ($Domain in $Forest.Domains) {
            foreach ($Section in $ADSectionsDomain) {
                $WordDocument = New-ADDocumentBlock `
                    -WordDocument $WordDocument `
                    -Section $Document.DocumentAD.Sections.SectionDomain.$Section `
                    -Forest $Forest `
                    -Domain $Domain `
                    -Excel $ExcelDocument `
                    -SectionName $Section
            }
        }
        ### End Sections

        ### Ending WORD
        $FilePath = Save-WordDocument -WordDocument $WordDocument `
            -Language $Document.Configuration.Prettify.Language `
            -FilePath $Document.DocumentAD.FilePathWord `
            -Supress $True `
            -OpenDocument:$Document.Configuration.Options.OpenDocument

        ### Ending EXCEL
        if ($Document.DocumentAD.ExportExcel) {
            Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $Document.DocumentAD.FilePathExcel -OpenWorkBook:$Document.Configuration.Options.OpenWorkbook
        }
    }
}
