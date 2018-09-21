function Start-DocumentationO365 {
    param(
        $Document
    )

    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Document.DocumentO365.Configuration

    if ($CheckCredentials) {
        if ($Document.DocumentO365.Configuration.PasswordFromFile) {
            if (Test-Path $Document.DocumentO365.Configuration.PasswordFromFile) {
                $Password = Get-Content $Document.DocumentO365.Configuration.PasswordFromFile
            }
        } else {
            $Password = $Document.DocumentO365.Configuration.Password
        }

        $Session = Connect-Exchange -SessionName $Document.DocumentO365.Configuration.ExchangeSessionName `
            -ConnectionURI $Document.DocumentO365.Configuration.ExchangeURI `
            -Authentication $Document.DocumentO365.Configuration.ExchangeAuthentication `
            -Username $Document.DocumentO365.Configuration.Username `
            -Password $Password `
            -AsSecure:$Document.DocumentO365.Configuration.PasswordAsSecure

        Import-PSSession -Session $Session -AllowClobber -DisableNameChecking

        Connect-Azure -SessionName $Document.DocumentO365.Configuration.AzureSessionName `


        $CheckAvailabilityCommands = Test-AvailabilityCommands -Commands 'Get-MailContact', 'Get-CalendarProcessing'
        if ($CheckAvailabilityCommands -notcontains $false) {
            $TypesRequired = Get-TypesRequired -Sections $Document.DocumentO365.Sections
            $DataSections = Get-ObjectKeys -Object $Document.DocumentO365.Sections

            $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
            $DataInformation = Get-WinExchangeInformation -TypesRequired $TypesRequired
            $TimeDataOnly.Stop()

            $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
            ### Starting WORD
            if ($Document.DocumentO365.ExportWord) {
                $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentO365.FilePathWord
            }
            if ($Document.DocumentO365.ExportExcel) {
                $ExcelDocument = New-ExcelDocument
            }
            ### Start Sections
            foreach ($Section in $DataSections) {
                $WordDocument = New-DataBlock `
                    -WordDocument $WordDocument `
                    -Section $Document.DocumentO365.Sections.$Section `
                    -Forest $DataInformation `
                    -Excel $ExcelDocument `
                    -SectionName $Section `
                    -Sql $Document.DocumentO365.ExportSQL
            }
            ### End Sections

            ### Ending WORD
            if ($Document.DocumentO365.ExportWord) {
                $FilePath = Save-WordDocument -WordDocument $WordDocument -Language $Document.Configuration.Prettify.Language -FilePath $Document.DocumentO365.FilePathWord -Supress $True -OpenDocument:$Document.Configuration.Options.OpenDocument
            }
            ### Ending EXCEL
            if ($Document.DocumentO365.ExportExcel) {
                $ExcelData = Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $Document.DocumentO365.FilePathExcel -OpenWorkBook:$Document.Configuration.Options.OpenExcel
            }
            $TimeDocuments.Stop()
            $TimeTotal.Stop()
            Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
            Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
            Write-Verbose "Time total: $($TimeTotal.Elapsed)"
        }
    }
}