function Start-DocumentationAWS {
    [CmdletBinding()]
    param(
        $Document
    )
    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Document.DocumentAWS.Configuration
    if ($CheckCredentials) {
        $TypesRequired = Get-TypesRequired -Sections $Document.DocumentAWS.Sections
        $DataSections = Get-ObjectKeys -Object $Document.DocumentAWS.Sections

        $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
        $DataInformation = Get-AWSInformation -TypesRequired $TypesRequired -AWSAccessKey $Document.DocumentAWS.Configuration.AWSAccessKey -AWSSecretKey $Document.DocumentAWS.Configuration.AWSSecretKey -AWSRegion $Document.DocumentAWS.Configuration.AWSRegion
        $TimeDataOnly.Stop()

        $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
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
        $TimeDocuments.Stop()
        $TimeTotal.Stop()
        Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
        Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
        Write-Verbose "Time total: $($TimeTotal.Elapsed)"
    }
}