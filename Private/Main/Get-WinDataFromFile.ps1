function Get-WinDataFromFile {
    [cmdletbinding()]
    param(
        [string] $FilePath,
        [string] $Type,
        [string] $FileType = 'XML'
    )
    try {
        if (Test-Path $FilePath) {
            if ($FileType -eq 'XML') {
                $Data = Import-Clixml -Path $FilePath -ErrorAction Stop
            } else {
                $File = Get-Content -Raw -Path $FilePath
                $Data = ConvertFrom-Json -InputObject $File
            }
        } else {
            Write-Warning "Couldn't load $FileType file from $FilePath for $Type data. File doesn't exists."
        }
    } catch {
        $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
        Write-Warning "Couldn't load $FileType file from $FilePath for $Type data. Error occured: $ErrorMessage"
    }
    return $Data
}