function Get-WinServiceData {
    [CmdletBinding()]
    param (
        [Object] $Credentials,
        [Object] $Service,
        [Object] $Type,
        [Object] $TypesRequired
    )
    if ($Service.Use) {
        if ($Service.OnlineMode) {
            switch ($Type) {
                'ActiveDirectory' {
                    # Prepare Data for Password Quality Checks
                    if ($Service.PasswordTests.Use) {
                        $PasswordClearText = $Service.PasswordTests.PasswordFilePathClearText
                    } else {
                        $PasswordClearText = ''
                    }
                    if ($Service.PasswordTests.UseHashDB) {
                        $PasswordHashes = $Service.PasswordTests.PasswordFilePathHash
                        if ($PasswordClearText -eq '') {
                            # creates temporary file to provide required data that is based on existance of this file
                            $TemporaryFile = New-TemporaryFile
                            'Passw0rd' | Out-File -FilePath $TemporaryFile.FullName
                            $PasswordClearText = $TemporaryFile.FullName
                        }
                    } else {
                        $PasswordHashes = ''
                    }

                    # Prepare Data AD
                    $CheckAvailabilityCommandsAD = Test-AvailabilityCommands -Commands 'Get-ADForest', 'Get-ADDomain', 'Get-ADRootDSE', 'Get-ADGroup', 'Get-ADUser', 'Get-ADComputer'
                    if ($CheckAvailabilityCommandsAD -contains $false) {
                        Write-Warning "Active Directory documentation can't be started as commands are unavailable. Check if you have Active Directory module available (part of RSAT) and try again."
                        return
                    }
                    if (-not (Test-ForestConnectivity)) {
                        Write-Warning 'Active DirectorNo connectivity to forest/domain.'
                        return
                    }
                    $DataInformation = Get-WinADForestInformation -TypesRequired $TypesRequired -PathToPasswords $PasswordClearText -PathToPasswordsHashes $PasswordHashes
                }
                'AWS' {
                    # Online mode
                    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Credentials -Verbose
                    if (-not $CheckCredentials) {
                        return
                    }
                    $DataInformation = Get-WinAWSInformation -TypesRequired $TypesRequired -AWSAccessKey $Credentials.AccessKey -AWSSecretKey $Credentials.SecretKey -AWSRegion $Credentials.Region

                }
                'Azure' {
                    # Check Credentials
                    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Credentials
                    if (-not $CheckCredentials) {
                        return
                    }
                    # Build Session
                    $Session = Connect-WinAzure -SessionName $Service.SessionName `
                        -Username $Credentials.Username `
                        -Password $Credentials.Password `
                        -AsSecure:$Credentials.PasswordAsSecure `
                        -FromFile:$Credentials.PasswordFromFile -Verbose

                    ## Gather Data
                    $DataInformation = Get-WinO365Azure -TypesRequired $TypesRequired
                    ## Plan for disconnect here
                }
                'AzureAD' {

                }
                'Exchange' {

                    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Document.DocumentExchange.Configuration -AllowEmptyKeys 'Username', 'Password'
                    if (-not $CheckCredentials) {
                        return
                    }
                    $Session = Connect-WinExchange -SessionName $Service.SessionName `
                        -ConnectionURI $Service.ConnectionURI `
                        -Authentication $Service.Authentication `
                        -Username $Credentials.Username `
                        -Password $Credentials.Password `
                        -AsSecure:$Credentials.PasswordAsSecure `
                        -FromFile:$Credentials.PasswordFromFile -Verbose

                    $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue' # weird but -Verbose:$false doesn't do anything
                    $ImportedSession = Import-PSSession -Session $Session -AllowClobber -DisableNameChecking -Verbose:$false
                    $VerbosePreference = $CurrentVerbosePreference

                    $CheckAvailabilityCommands = Test-AvailabilityCommands -Commands 'Get-ExchangeServer', 'Get-MailboxDatabase', 'Get-PublicFolderDatabase'
                    if ($CheckAvailabilityCommands -contains $false) {
                        return
                    }
                    ## Gather Data
                    $DataInformation = Get-WinExchangeInformation -TypesRequired $TypesRequired
                }
                'ExchangeOnline' {
                    # Check Credentials
                    $CheckCredentials = Test-ConfigurationCredentials -Configuration $Credentials
                    if (-not $CheckCredentials) {
                        return
                    }
                    # Build Session
                    $Session = Connect-WinExchange -SessionName $Service.SessionName `
                        -ConnectionURI $Service.ConnectionURI `
                        -Authentication $Service.Authentication `
                        -Username $Credentials.Username `
                        -Password $Credentials.Password `
                        -AsSecure:$Credentials.PasswordAsSecure `
                        -FromFile:$Credentials.PasswordFromFile -Verbose

                    # Failed connecting to session
                    if (-not $Session) {
                        return
                    }
                    # Import Session
                    $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue' # weird but -Verbose:$false doesn't do anything below
                    $ImportedSession = Import-PSSession -Session $Session -AllowClobber -DisableNameChecking -Prefix $Service.Prefix -Verbose:$false
                    $VerbosePreference = $CurrentVerbosePreference

                    ## Verify Connectivity
                    $CheckAvailabilityCommands = Test-AvailabilityCommands -Commands "Get-$($Service.Prefix)MailContact", "Get-$($Service.Prefix)CalendarProcessing"
                    if ($CheckAvailabilityCommands -contains $false) {
                        return
                    }
                    ## Gather Data
                    $DataInformation = Get-WinO365Exchange -TypesRequired $TypesRequired
                    ## Plan for disconnect here
                }
                'Teams' {

                }
                'SharePointOnline' {

                }
                'SkypeOnline' {

                }
            }
            if ($Service.Export.Use) {
                $Time = Start-TimeLog
                if ($Service.Export.To -eq 'File' -or $Service.Export.To -eq 'Both') {
                    Save-WinDataToFile -Export $Service.Export.Use -FilePath $Service.Export.FilePath -Data $DataInformation -Type $Type -IsOffline:$false -FileType 'XML'
                    $TimeSummary = Stop-TimeLog -Time $Time -Option OneLiner
                    Write-Verbose "Saving data for $Type to file $($Service.Export.FilePath) took: $TimeSummary"
                }
                if ($Service.Export.To -eq 'Folder' -or $Service.Export.To -eq 'Both') {
                    $Time = Start-TimeLog
                    Save-WinDataToFileInChunks -Export $Service.Export.Use -FolderPath $Service.Export.FolderPath -Data $DataInformation -Type $Type -IsOffline:$false -FileType 'XML'
                    $TimeSummary = Stop-TimeLog -Time $Time -Option OneLiner
                    Write-Verbose "Saving data for $Type to folder $($Service.Export.FolderPath) took: $TimeSummary"
                }
            }
            return $DataInformation
        } else {
            if ($Service.Import.Use) {
            $Time = Start-TimeLog
            if ($Service.Import.From -eq 'File') {
                Write-Verbose "Loading data for $Type in offline mode from XML File $($Service.Import.FilePath). Hang on..."
                $DataInformation = Get-WinDataFromFile -FilePath $Service.Import.FilePath -Type $Type -FileType 'XML'
            } elseif ($Service.Import.From -eq 'Folder') {
                Write-Verbose "Loading data for $Type in offline mode from XML File $($Service.Import.FilePath). Hang on..."
                $DataInformation = Get-WinDataFromFileInChunks -FolderPath $Service.Import.FolderPath -Type $Type -FileType 'XML'
            } else {
                Write-Warning "Wrong option for Import.Use. Only Folder/File is supported."
            }
            $TimeSummary = Stop-TimeLog -Time $Time -Option OneLiner
            Write-Verbose "Loading data for $Type in offline mode from file took $TimeSummary"
            return $DataInformation
            }

        }
    }
}