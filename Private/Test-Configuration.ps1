function Test-Configuration {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    [int] $ErrorCount = 0
    $Script:WriteParameters = $Document.Configuration.DisplayConsole


    $Keys = Get-ObjectKeys -Object $Document -Ignore 'Configuration'
    foreach ($Key in $Keys) {
        $ErrorCount += Test-File -File $Document.$Key.FilePathWord -FileName 'FilePathWord' -Skip:(-not $Document.$Key.ExportWord)
        $ErrorCount += Test-File -File $Document.$Key.FilePathExcel -FileName 'FilePathExcel' -Skip:(-not $Document.$Key.ExportExcel)
    }
    if ($ErrorCount -ne 0) {
        Exit
    }
}