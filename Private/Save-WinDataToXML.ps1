function Save-WinDataToXML {
    param(
        [bool] $Export,
        [string] $Type,
        $Data,
        [string] $FilePath,
        [switch] $IsOffline
    )
    if ($IsOffline) {
        # This means data is loaded from xml so it doesn't need to be resaved to XML
        return
    }
    if ($Export) {
        if ($FilePath) {
            Write-Verbose "Exporting $Type data to XML to path $FilePath"
            try {
                $Data | Export-Clixml -Path $FilePath -ErrorAction Stop
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                Write-Warning "Couldn't save XML file to $FilePath for $Type data. Error occured: $ErrorMessage"
            }
        }
    }
}