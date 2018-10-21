function Get-WinDataFromXML {
    [cmdletbinding()]
    param(
        $FilePath,
        [string] $Type
    )
    try {
        if (Test-Path $FilePath) {
            $Data = Import-Clixml -Path $FilePath -ErrorAction Stop
        } else {
            Write-Warning "Couldn't load XML file from $FilePath for $Type data. File doesn't exists."
        }
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Warning "Couldn't load XML file from $FilePath for $Type data. Error occured: $ErrorMessage"
    }
    return $Data
}