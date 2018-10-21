function Save-WinDataToXML {
    [cmdletbinding()]
    param(
        [nullable[bool]] $Export,
        [string] $Type,
        $Data,
        [string] $FilePath,
        [switch] $IsOffline
    )
    if ($IsOffline) {
        # This means data is loaded from xml so it doesn't need to be resaved to XML
        Write-Verbose "Save-WinDataToXML - Exporting $Type data to XML to path $FilePath skipped. Running in offline mode."
        return
    }
    if ($Export) {
        if ($FilePath) {
            Write-Verbose "Save-WinDataToXML - Exporting $Type data to XML to path $FilePath"
            try {
                $Data | Export-Clixml -Path $FilePath -ErrorAction Stop
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                Write-Warning "Couldn't save XML file to $FilePath for $Type data. Error occured: $ErrorMessage"
            }
        }
    }
}