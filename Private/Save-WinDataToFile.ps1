function Save-WinDataToFile {
    [cmdletbinding()]
    param(
        [nullable[bool]] $Export,
        [string] $Type,
        [Object] $Data,
        [string] $FilePath,
        [switch] $IsOffline,
        [string] $FileType = 'XML'
    )
    if ($IsOffline) {
        # This means data is loaded from xml so it doesn't need to be resaved to XML
        Write-Verbose "Save-WinDataToFile - Exporting $Type data to $FileType to path $FilePath skipped. Running in offline mode."
        return
    }
    if ($Export) {
        if ($FilePath) {
            $Split = Split-Path -Path $FilePath
            if (-not (Test-Path -Path $Split)) {
                # Creates directory path if it doesn't exits
                New-Item -ItemType Directory -Force -Path $Split > $null
            }
            Write-Verbose "Save-WinDataToFile - Exporting $Type data to $FileType to path $FilePath"
            if ($FileType -eq 'XML') {
                try {
                    $Data | Export-Clixml -Path $FilePath -ErrorAction Stop -Encoding UTF8
                } catch {
                    $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                    Write-Warning "Couldn't save $FileType file to $FilePath for $Type data. Error occured: $ErrorMessage"
                }
            } else {
                try {
                    $Data | ConvertTo-Json -ErrorAction Stop  | Add-Content -Path $FilePath -Encoding UTF8 -ErrorAction Stop
                } catch {
                    $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                    Write-Warning "Couldn't save $FileType file to $FilePath for $Type data. Error occured: $ErrorMessage"
                }
            }
        }
    }
}