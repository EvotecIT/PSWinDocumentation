function Start-DocumentationO365 {
    [CmdletBinding()]
    param(
        $Document
    )

    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Document.DocumentOffice365.Configuration

    if ($CheckCredentials) {
        if ($Document.DocumentOffice365.Configuration.O365ExchangeUse) {
            $Session = Connect-Exchange -SessionName $Document.DocumentOffice365.Configuration.O365ExchangeSessionName `
                -ConnectionURI $Document.DocumentOffice365.Configuration.O365ExchangeURI `
                -Authentication $Document.DocumentOffice365.Configuration.O365ExchangeAuthentication `
                -Username $Document.DocumentOffice365.Configuration.O365Username `
                -Password $Document.DocumentOffice365.Configuration.O365Password `
                -AsSecure:$Document.DocumentOffice365.Configuration.O365PasswordAsSecure `
                -FromFile:$Document.DocumentOffice365.Configuration.O365PasswordFromFile -Verbose

            # Failed connecting to exchange
            if (-not $Session) { return }

            $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue' # weird but -Verbose:$false doesn't do anything below
            $ImportedSession = Import-PSSession -Session $Session -AllowClobber -DisableNameChecking -Prefix 'O365' -Verbose:$false
            $VerbosePreference = $CurrentVerbosePreference
        }

        if ($Document.DocumentOffice365.Configuration.O365AzureADUse) {
            $SessionAzure = Connect-Azure -SessionName $Document.DocumentOffice365.Configuration.O365AzureSessionName `
                -Username $Document.DocumentOffice365.Configuration.O365Username `
                -Password $Document.DocumentOffice365.Configuration.O365Password `
                -AsSecure:$Document.DocumentOffice365.Configuration.O365PasswordAsSecure `
                -FromFile:$Document.DocumentOffice365.Configuration.O365PasswordFromFile -Verbose
        }

        $TypesRequired = Get-TypesRequired -Sections $Document.DocumentOffice365.Sections
        $DataSections = Get-ObjectKeys -Object $Document.DocumentOffice365.Sections

        $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start

        if ($Document.DocumentOffice365.Configuration.O365ExchangeUse) {
            $CheckAvailabilityCommands = Test-AvailabilityCommands -Commands 'Get-O365MailContact', 'Get-O365CalendarProcessing'
            if ($CheckAvailabilityCommands -notcontains $false) {
                $DataInformationO365 = Get-WinO365Exchange -TypesRequired $TypesRequired
            }
        } else {
            $DataInformationO365 = $null
        }
        if ($Document.DocumentOffice365.Configuration.O365AzureADUse) {
            $DataInformationAzure = Get-WinO365Azure -TypesRequired $TypesRequired
        } else {
            $DataInformationAzure = $null
        }
        $TimeDataOnly.Stop()

        $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
        ### Starting WORD
        if ($Document.DocumentOffice365.ExportWord) {
            $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentOffice365.FilePathWord
        }
        if ($Document.DocumentOffice365.ExportExcel) {
            $ExcelDocument = New-ExcelDocument
        }

        ### Start Sections
        foreach ($DataInformation in $($DataInformationO365 + $DataInformationAzure)) {
            foreach ($Section in $DataSections) {
                $WordDocument = New-DataBlock `
                    -WordDocument $WordDocument `
                    -Section $Document.DocumentOffice365.Sections.$Section `
                    -Forest $DataInformation `
                    -Excel $ExcelDocument `
                    -SectionName $Section `
                    -Sql $Document.DocumentOffice365.ExportSQL
            }
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
        $TimeDocuments.Stop()
        $TimeTotal.Stop()
        Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
        Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
        Write-Verbose "Time total: $($TimeTotal.Elapsed)"
    }
}