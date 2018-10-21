function Start-DocumentationAWS {
    [CmdletBinding()]
    param(
        $Document
    )
    $DataSections = Get-ObjectKeys -Object $Document.DocumentAWS.Sections
    $TypesRequired = Get-TypesRequired -Sections $Document.DocumentAWS.Sections

    ### Start AWS Data
    $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    if ($Document.DocumentAWS.Configuration.OfflineMode.Use) {
        # Offline mode
        if ($Document.DocumentAWS.ExportXML) {
            Write-Warning "You can't run AWS Documentation in 'offline mode' with 'ExportXML' set to true. Please turn off one of the options."
            return
        } else {
            $DataInformation = Get-WinDataFromXML -FilePath $Document.DocumentAWS.Configuration.OfflineMode.XMLPath -Type [AWS]
        }
    } else {
        # Online mode
        $CheckCredentials = Test-ConfigurationCredentials -Configuration $Document.DocumentAWS.Configuration
        if ($CheckCredentials) {
            $DataInformation = Get-WinAWSInformation -TypesRequired $TypesRequired -AWSAccessKey $Document.DocumentAWS.Services.AWS.AWSAccessKey -AWSSecretKey $Document.DocumentAWS.Services.AWS.AWSSecretKey -AWSRegion $Document.DocumentAWS.Services.AWS.AWSRegion
        }
    }
    $TimeDataOnly.Stop()
    ### End AWS Data
    $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    # Saves data to XML is required - skipped when Offline mode is on
    if ($DataInformation) {
        Save-WinDataToXML -Export $Document.DocumentAWS.ExportXML -FilePath $Document.DocumentAWS.FilePathXML -Data $DataInformationAD -Type [AWS] -IsOffline:$Document.DocumentAWS.Configuration.OfflineMode.Use

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
    $TimeTotal.Stop()
    Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
    Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
    Write-Verbose "Time total: $($TimeTotal.Elapsed)"

}