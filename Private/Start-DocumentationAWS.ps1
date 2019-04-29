function Start-DocumentationAWS {
    [CmdletBinding()]
    param(
        $Document
    )
    $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start

    $DataSections = Get-ObjectKeys -Object $Document.DocumentAWS.Sections
    $TypesRequired = Get-TypesRequired -Sections $Document.DocumentAWS.Sections

    $DataInformation = Get-WinServiceData -Credentials $Document.DocumentAWS.Services.Amazon.Credentials `
        -Service $Document.DocumentAWS.Services.Amazon.AWS `
        -TypesRequired $TypesRequired `
        -Type 'AWS'

    $TimeDataOnly.Stop()

    $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    # Saves data to XML is required - skipped when Offline mode is on
    if ($DataInformation.Count -gt 0) {
        ### Starting WORD
        if ($Document.DocumentAWS.ExportWord) {
            $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentAWS.FilePathWord
        }
        if ($Document.DocumentAWS.ExportExcel) {
            $ExcelDocument = New-ExcelDocument
        }
        ### Start Sections
        foreach ($Section in $DataSections) {
            $WordDocument = New-DataBlock `
                -WordDocument $WordDocument `
                -Section $Document.DocumentAWS.Sections.$Section `
                -Forest $DataInformation `
                -Excel $ExcelDocument `
                -SectionName $Section `
                -Sql $Document.DocumentAWS.ExportSQL
        }
        ### End Sections

        ### Ending WORD
        if ($Document.DocumentAWS.ExportWord) {
            $FilePath = Save-WordDocument -WordDocument $WordDocument -Language $Document.Configuration.Prettify.Language -FilePath $Document.DocumentAWS.FilePathWord -Supress $True -OpenDocument:$Document.Configuration.Options.OpenDocument
        }
        ### Ending EXCEL
        if ($Document.DocumentAWS.ExportExcel) {
            $ExcelData = Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $Document.DocumentAWS.FilePathExcel -OpenWorkBook:$Document.Configuration.Options.OpenExcel
        }
    } else {
        Write-Warning "There was no data to process AWS documentation. Check configuration."
    }
    $TimeDocuments.Stop()
    Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
    Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
}