function Start-DocumentationO365 {
    [CmdletBinding()]
    param(
        [Object] $Document
    )
    $TypesRequired = Get-TypesRequired -Sections $Document.DocumentOffice365.Sections
    $DataSections = Get-ObjectKeys -Object $Document.DocumentOffice365.Sections

    $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start


    $DataAzure = Get-WinServiceData -Credentials $Document.DocumentOffice365.Services.Office365.Credentials `
        -Service $Document.DocumentOffice365.Services.Office365.Azure `
        -TypesRequired $TypesRequired `
        -Type 'Azure'
    $DataExchangeOnline = Get-WinServiceData -Credentials $Document.DocumentOffice365.Services.Office365.Credentials `
        -Service $Document.DocumentOffice365.Services.Office365.ExchangeOnline `
        -TypesRequired $TypesRequired `
        -Type 'ExchangeOnline'

    $DataInformation = [ordered]@{}
    if ($null -ne $DataAzure -and $DataExchangeOnline.Count -gt 0) {
        $DataInformation += $DataAzure
    }
    if ($null -ne $DataExchangeOnline -and $DataExchangeOnline.Count -gt 0) {
        $DataInformation += $DataExchangeOnline
    }
    $TimeDataOnly.Stop()


    $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    ### Starting WORD

    if ($DataInformation.Count -gt 0) {
        if ($Document.DocumentOffice365.ExportWord) {
            $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentOffice365.FilePathWord
        }
        if ($Document.DocumentOffice365.ExportExcel) {
            $ExcelDocument = New-ExcelDocument
        }

        ### Start Sections
        foreach ($Section in $DataSections) {
            $WordDocument = New-DataBlock `
                -WordDocument $WordDocument `
                -Section $Document.DocumentOffice365.Sections.$Section `
                -Forest $DataInformation `
                -Excel $ExcelDocument `
                -SectionName $Section `
                -Sql $Document.DocumentOffice365.ExportSQL
        }
        ### End Sections

        ### Ending WORD
        if ($Document.DocumentOffice365.ExportWord) {
            $FilePath = Save-WordDocument -WordDocument $WordDocument -Language $Document.Configuration.Prettify.Language -FilePath $Document.DocumentOffice365.FilePathWord -Supress $True -OpenDocument:$Document.Configuration.Options.OpenDocument
        }
        ### Ending EXCEL
        if ($Document.DocumentOffice365.ExportExcel) {
            $ExcelData = Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $Document.DocumentOffice365.FilePathExcel -OpenWorkBook:$Document.Configuration.Options.OpenExcel
        }
    } else {
        Write-Warning "There was no data to process Office 365 documentation. Check configuration."
    }

    $TimeDocuments.Stop()
    $TimeTotal.Stop()
    Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
    Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
    Write-Verbose "Time total: $($TimeTotal.Elapsed)"
}