function Get-WinDataFromXML {
    param(
        $FilePath,
        [string] $Type
    )
    try {
        $Data = Import-Clixml -Path $FilePath -ErrorAction Stop
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Warning "Couldn't load XML file from $FilePath for $Type data. Error occured: $ErrorMessage"
    }
    return $Data
}