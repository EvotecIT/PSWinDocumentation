function Start-DocumentationExchange {
    [CmdletBinding()]
    param(
        $Document
    )
    $DataSections = Get-ObjectKeys -Object $Document.DocumentExchange.Sections
    $TypesRequired = Get-TypesRequired -Sections $Document.DocumentExchange.Sections

    $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start

    ### Start Exchange Data
    if ($Document.DocumentExchange.Configuration.OfflineMode.Use) {
        # Offline mode
        if ($Document.DocumentExchange.ExportXML) {
            Write-Warning "You can't run Microsoft Exchange Documentation in 'offline mode' with 'ExportXML' set to true. Please turn off one of the options."
            return
        } else {
            $DataInformation = Get-WinDataFromXML -FilePath $Document.DocumentExchange.Configuration.OfflineMode.XMLPath -Type [Exchange]
        }
    } else {
        # Online mode
        $CheckCredentials = Test-ConfigurationCredentials -Configuration $Document.DocumentExchange.Configuration -AllowEmptyKeys 'Username', 'Password'
        if ($CheckCredentials) {
            if ($Document.DocumentExchange.Configuration.PasswordFromFile) {
                if (Test-Path $Document.DocumentExchange.Configuration.PasswordFromFile) {
                    $Password = Get-Content $Document.DocumentExchange.Configuration.PasswordFromFile
                }
            } else {
                $Password = $Document.DocumentExchange.Configuration.Password
            }
            $Session = Connect-WinExchange -SessionName $Document.DocumentExchange.Configuration.ExchangeSessionName `
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
                $DataInformation = Get-WinExchangeInformation -TypesRequired $TypesRequired
            }
        }
    }
    $TimeDataOnly.Stop()
    # End Exchange Data
    $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start

    if ($DataInformation) {
        Save-WinDataToXML -Export $Document.DocumentExchange.ExportXML -FilePath $Document.DocumentExchange.FilePathXML -Data $DataInformationAD -Type [Exchange] -IsOffline:$Document.DocumentExchange.Configuration.OfflineMode.Use

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
    } else {
        Write-Warning "There was no data to process Exchange documentation. Check configuration."
    }

    $TimeDocuments.Stop()
    $TimeTotal.Stop()
    Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
    Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
    Write-Verbose "Time total: $($TimeTotal.Elapsed)"


}