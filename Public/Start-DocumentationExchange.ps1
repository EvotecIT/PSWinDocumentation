function Start-DocumentationExchange {
    [CmdletBinding()]
    param(
        $Document
    )

    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Document.DocumentExchange.Configuration -AllowEmptyKeys 'Username', 'Password'

    if ($CheckCredentials) {

        if ($Document.DocumentExchange.Configuration.PasswordFromFile) {
            if (Test-Path $Document.DocumentExchange.Configuration.PasswordFromFile) {
                $Password = Get-Content $Document.DocumentExchange.Configuration.PasswordFromFile
            }
        } else {
            $Password = $Document.DocumentExchange.Configuration.Password
        }

        $Session = Connect-Exchange -SessionName $Document.DocumentExchange.Configuration.ExchangeSessionName `
            -ConnectionURI $Document.DocumentExchange.Configuration.ExchangeURI `
            -Authentication $Document.DocumentExchange.Configuration.ExchangeAuthentication `
            -Username $Document.DocumentExchange.Configuration.Username `
            -Password $Password `
            -AsSecure:$Document.DocumentExchange.Configuration.PasswordAsSecure

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue' # weird but -Verbose:$false doesn't do anything
        $ImportedSession = Import-PSSession -Session $Session -AllowClobber -DisableNameChecking -Verbose:$false
        $VerbosePreference = $CurrentVerbosePreference

        $CheckAvailabilityCommands = Test-AvailabilityCommands -Commands 'Get-ExchangeServer', 'Get-MailboxDatabase', 'Get-PublicFolderDatabase'
        if ($CheckAvailabilityCommands -notcontains $false) {
            $TypesRequired = Get-TypesRequired -Sections $Document.DocumentExchange.Sections
            $DataSections = Get-ObjectKeys -Object $Document.DocumentExchange.Sections

            $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
            $DataInformation = Get-WinExchangeInformation -TypesRequired $TypesRequired
            $TimeDataOnly.Stop()

            $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
            ### Starting WORD
            if ($Document.DocumentExchange.ExportWord) {
                $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentExchange.FilePathWord
            }
            if ($Document.DocumentExchange.ExportExcel) {
                $ExcelDocument = New-ExcelDocument
            }
            ### Start Sections
            foreach ($Section in $DataSections) {
                $WordDocument = New-DataBlock `
                    -WordDocument $WordDocument `
                    -Section $Document.DocumentExchange.Sections.$Section `
                    -Forest $DataInformation `
                    -Excel $ExcelDocument `
                    -SectionName $Section `
                    -Sql $Document.DocumentExchange.ExportSQL
            }
            ### End Sections

            ### Ending WORD
            if ($Document.DocumentExchange.ExportWord) {
                $FilePath = Save-WordDocument -WordDocument $WordDocument -Language $Document.Configuration.Prettify.Language -FilePath $Document.DocumentExchange.FilePathWord -Supress $True -OpenDocument:$Document.Configuration.Options.OpenDocument
            }
            ### Ending EXCEL
            if ($Document.DocumentExchange.ExportExcel) {
                $ExcelData = Save-ExcelDocument -ExcelDocument $ExcelDocument -FilePath $Document.DocumentExchange.FilePathExcel -OpenWorkBook:$Document.Configuration.Options.OpenExcel
            }
            $TimeDocuments.Stop()
            $TimeTotal.Stop()
            Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
            Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
            Write-Verbose "Time total: $($TimeTotal.Elapsed)"
        }
    }
}