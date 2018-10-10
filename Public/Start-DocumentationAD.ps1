function Start-DocumentationAD {
    [CmdletBinding()]
    param(
        $Document
    )
    $TypesRequired = Get-TypesRequired -Sections $Document.DocumentAD.Sections.SectionForest, $Document.DocumentAD.Sections.SectionDomain
    $ADSectionsForest = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionForest
    $ADSectionsDomain = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionDomain

    $ADConfiguration = $Document.DocumentAD.Configuration
    if ($ADConfiguration.PasswordTests.Use) {
        $PasswordClearText = $ADConfiguration.PasswordTests.PasswordFilePathClearText
    } else {
        $PasswordClearText = ''
    }
    if ($ADConfiguration.PasswordTests.UseHashDB) {
        $PasswordHashes = $ADConfiguration.PasswordTests.PasswordFilePathHash
    } else {
        $PasswordHashes = ''
    }

    $TimeDataOnly = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    if ($Document.DocumentAD.Configuration.OfflineMode.Use) {
        # Offline mode
        $DataInformationAD = Get-WinDataFromXML -FilePath $Document.DocumentAD.Configuration.OfflineMode.XMLPath -Type [ActiveDirectory]
    } else {
        # Online mode
        $CheckAvailabilityCommandsAD = Test-AvailabilityCommands -Commands 'Get-ADForest', 'Get-ADDomain', 'Get-ADRootDSE', 'Get-ADGroup', 'Get-ADUser', 'Get-ADComputer'
        if ($CheckAvailabilityCommandsAD -notcontains $false) {
            Test-ForestConnectivity
            $DataInformationAD = Get-WinADForestInformation -TypesRequired $TypesRequired -PathToPasswords $PasswordClearText -PathToPasswordsHashes $PasswordHashes

        } else {
            Write-Warning "Active Directory documentation can't be started as commands are unavailable. Check if you have Active Directory module available (part of RSAT) and try again."
            return
        }
    }
    $TimeDataOnly.Stop()
    $TimeDocuments = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start

    # Saves data to XML is required - skipped when Offline mode is on
    Save-WinDataToXML -Export $Document.DocumentAD.ExportXML -FilePath $Document.DocumentAD.FilePathXML -Data $DataInformationAD -Type [ActiveDirectory] -IsOffline:$Document.DocumentAD.Configuration.OfflineMode.Use

    if ($Document.DocumentAD.ExportExcel -or $Document.DocumentAD.ExportWord -or $Document.DocumentAD.ExportSQL) {

        ### Starting WORD
        if ($Document.DocumentAD.ExportWord) {
            $WordDocument = Get-DocumentPath -Document $Document -FinalDocumentLocation $Document.DocumentAD.FilePathWord
        }
        if ($Document.DocumentAD.ExportExcel) {
            $ExcelDocument = New-ExcelDocument
        }
        ### Start Sections
        foreach ($DataInformation in $DataInformationAD) {
            foreach ($Section in $ADSectionsForest) {
                $WordDocument = New-DataBlock `
                    -WordDocument $WordDocument `
                    -Section $Document.DocumentAD.Sections.SectionForest.$Section `
                    -Object $DataInformationAD `
                    -Excel $ExcelDocument `
                    -SectionName $Section `
                    -Sql $Document.DocumentAD.ExportSQL
            }
            foreach ($Domain in $DataInformationAD.Domains) {
                foreach ($Section in $ADSectionsDomain) {
                    $WordDocument = New-DataBlock `
                        -WordDocument $WordDocument `
                        -Section $Document.DocumentAD.Sections.SectionDomain.$Section `
                        -Object $DataInformationAD `
                        -Domain $Domain `
                        -Excel $ExcelDocument `
                        -SectionName $Section `
                        -Sql $Document.DocumentAD.ExportSQL
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
    $TimeDocuments.Stop()
    $TimeTotal.Stop()
    Write-Verbose "Time to gather data: $($TimeDataOnly.Elapsed)"
    Write-Verbose "Time to create documents: $($TimeDocuments.Elapsed)"
    Write-Verbose "Time total: $($TimeTotal.Elapsed)"
}