function Start-Documentation {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    $TimeTotal = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    Test-Configuration -Document $Document

    if ($Document.DocumentAD.Enable) {
        Test-ModuleAvailability
        Test-ForestConnectivity
        $TypesRequired = Get-TypesRequired -Sections $Document.DocumentAD.Sections.SectionForest, $Document.DocumentAD.Sections.SectionDomain

        $ADSectionsForest = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionForest
        $ADSectionsDomain = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionDomain

        $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
        $Forest = Get-WinADForestInformation -TypesRequired $TypesRequired
        $TimeDataOnly.Stop()
        $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
        ### Starting WORD
        if ($Document.DocumentAD.ExportWord) {
            $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentAD.FilePathWord
        }
        if ($Document.DocumentAD.ExportExcel) {
            $ExcelDocument = New-ExcelDocument
        }
        ### Start Sections
        foreach ($Section in $ADSectionsForest) {
            $WordDocument = New-DataBlock `
                -WordDocument $WordDocument `
                -Section $Document.DocumentAD.Sections.SectionForest.$Section `
                -Forest $Forest `
                -Excel $ExcelDocument `
                -SectionName $Section `
                -Sql $Document.DocumentAD.ExportSQL
        }
        foreach ($Domain in $Forest.Domains) {
            foreach ($Section in $ADSectionsDomain) {
                $WordDocument = New-DataBlock `
                    -WordDocument $WordDocument `
                    -Section $Document.DocumentAD.Sections.SectionDomain.$Section `
                    -Object $Forest `
                    -Domain $Domain `
                    -Excel $ExcelDocument `
                    -SectionName $Section `
                    -Sql $Document.DocumentAD.ExportSQL
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
        $TimeDocuments.Stop()
        $TimeTotal.Stop()
        Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
        Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
        Write-Verbose "Time total: $($TimeTotal.Elapsed)"
    }
    if ($Document.DocumentAWS.Enable) {
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
    if ($Document.DocumentExchange.Enable) {
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
                -AsSecure:$Document.DocumentExchange.Configuration.PasswordAsSecure `

            Import-PSSession -Session $Session -AllowClobber -DisableNameChecking

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
}