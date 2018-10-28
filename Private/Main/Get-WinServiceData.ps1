function Get-WinServiceData {
    [CmdletBinding()]
    param (
        [Object] $Credentials,
        [Object] $Service,
        [Object] $Type,
        [Object] $TypesRequired
    )

    Connect-WinService -Type $Type `
        -Credentials $Credentials `
        -Service $Service

    if ($Service.Use) {
        if ($Service.OnlineMode) {
            switch ($Type) {
                'ActiveDirectory' {
                    $DataInformation = Get-WinADForestInformation -TypesRequired $TypesRequired -PathToPasswords $PasswordClearText -PathToPasswordsHashes $PasswordHashes
                }
                'AWS' {
                    $DataInformation = Get-WinAWSInformation -TypesRequired $TypesRequired -AWSAccessKey $Credentials.AccessKey -AWSSecretKey $Credentials.SecretKey -AWSRegion $Credentials.Region
                }
                'Azure' {
                    $DataInformation = Get-WinO365Azure -TypesRequired $TypesRequired
                }
                'AzureAD' {

                }
                'Exchange' {
                    $DataInformation = Get-WinExchangeInformation -TypesRequired $TypesRequired
                }
                'ExchangeOnline' {
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
                    Write-Verbose "Loading data for $Type in offline mode from XML File $($Service.Import.Path). Hang on..."
                    $DataInformation = Get-WinDataFromFile -FilePath $Service.Import.Path -Type $Type -FileType 'XML'
                } elseif ($Service.Import.From -eq 'Folder') {
                    Write-Verbose "Loading data for $Type in offline mode from XML File $($Service.Import.Path). Hang on..."
                    $DataInformation = Get-WinDataFromFileInChunks -FolderPath $Service.Import.Path -Type $Type -FileType 'XML'
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