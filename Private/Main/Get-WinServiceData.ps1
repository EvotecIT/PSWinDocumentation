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

                    # Failed connecting to session
                    #if (-not $Session) {
                    #    return
                    #}

                    ## Gather Data
                    $DataInformation = Get-WinO365Azure -TypesRequired $TypesRequired
                    if ($Service.ExportXML) {
                        Save-WinDataToXML -Export $Service.ExportXML -FilePath $Service.ExportXMLPath -Data $DataInformation -Type [O365] -IsOffline:(-not $Service.OnlineMode)
                    }
                    ## Plan for disconnect here

                    ## Return Data
                    return $DataInformation
                }
                'AzureAD' {

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
                    if ($Service.ExportXML) {
                        Save-WinDataToXML -Export $Service.ExportXML -FilePath $Service.ExportXMLPath -Data $DataInformation -Type [O365] -IsOffline:(-not $Service.OnlineMode)
                    }

                    ## Plan for disconnect here

                    ## Return Data
                    return $DataInformation

                }
                'Teams' {

                }
                'SharePointOnline' {

                }
                'SkypeOnline' {

                }
            }
        } else {
            $DataInformation = Get-WinDataFromXML -FilePath $Service.XMLPath -Type $Type
            return $DataInformation
        }
    }
}