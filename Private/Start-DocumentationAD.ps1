function Start-DocumentationAD {
    [CmdletBinding()]
    param(
        [System.Collections.IDictionary] $Document
    )
    $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    $TypesRequired = Get-TypesRequired -Sections $Document.DocumentAD.Sections.SectionForest, $Document.DocumentAD.Sections.SectionDomain
    $DataInformationAD = Get-WinServiceData -Credentials $Document.DocumentAD.Services.OnPremises.Credentials `
        -Service $Document.DocumentAD.Services.OnPremises.ActiveDirectory `
        -TypesRequired $TypesRequired `
        -Type 'ActiveDirectory'

    $TimeDataOnly.Stop()
    $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start

    # Saves data to XML is required - skipped when Offline mode is on
    #if ($DataInformationAD) {
    if ($Document.DocumentAD.ExportExcel -or $Document.DocumentAD.ExportWord -or $Document.DocumentAD.ExportSQL) {

        ### Starting WORD
        if ($Document.DocumentAD.ExportWord) {
            $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentAD.FilePathWord
        }
        if ($Document.DocumentAD.ExportExcel) {
            $ExcelDocument = New-ExcelDocument
        }

        $ADSectionsForest = ($Document.DocumentAD.Sections.SectionForest).Keys
        $ADSectionsDomain = ($Document.DocumentAD.Sections.SectionDomain).Keys
        ### Start Sections
        foreach ($DataInformation in $DataInformationAD) {
            foreach ($Section in $ADSectionsForest) {
                if ($WordDocument) {
                    $WordDocument = New-DataBlock `
                        -WordDocument $WordDocument `
                        -Section $Document.DocumentAD.Sections.SectionForest.$Section `
                        -Object $DataInformationAD `
                        -Excel $ExcelDocument `
                        -SectionName $Section `
                        -Sql $Document.DocumentAD.ExportSQL -ExportWord $Document.DocumentAD.ExportWord
                } else {
                    New-DataBlock `
                        -Section $Document.DocumentAD.Sections.SectionForest.$Section `
                        -Object $DataInformationAD `
                        -Excel $ExcelDocument `
                        -SectionName $Section `
                        -Sql $Document.DocumentAD.ExportSQL -ExportWord $Document.DocumentAD.ExportWord
                }
            }
            foreach ($Domain in $DataInformationAD.FoundDomains.Keys) {
                foreach ($Section in $ADSectionsDomain) {
                    if ($WordDocument) {
                        $WordDocument = New-DataBlock `
                            -WordDocument $WordDocument `
                            -Section $Document.DocumentAD.Sections.SectionDomain.$Section `
                            -Object $DataInformationAD `
                            -Domain $Domain `
                            -Excel $ExcelDocument `
                            -SectionName $Section `
                            -Sql $Document.DocumentAD.ExportSQL -ExportWord $Document.DocumentAD.ExportWord
                    } else {
                        New-DataBlock `
                            -Section $Document.DocumentAD.Sections.SectionDomain.$Section `
                            -Object $DataInformationAD `
                            -Domain $Domain `
                            -Excel $ExcelDocument `
                            -SectionName $Section `
                            -Sql $Document.DocumentAD.ExportSQL -ExportWord $Document.DocumentAD.ExportWord
                    }
                }
            }
        }
        ### End Sections

        ### Ending WORD
        if ($Document.DocumentAD.ExportWord) {
            $FilePath = Save-WordDocument -WordDocument $WordDocument `
                -Language $Document.Configuration.Prettify.Language `
                -FilePath $Document.DocumentAD.FilePathWord `
                -Supress $True `
                -OpenDocument:$Document.Configuration.Options.OpenDocument
        }
        ### Ending EXCEL
        if ($Document.DocumentAD.ExportExcel) {
            $ExcelData = Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $Document.DocumentAD.FilePathExcel -OpenWorkBook:$Document.Configuration.Options.OpenExcel
        }
    }

    #}
    $TimeDocuments.Stop()
    Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
    Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
}