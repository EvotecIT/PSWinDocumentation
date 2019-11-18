function Get-TypesRequired {
    [CmdletBinding()]
    param (
        [System.Collections.IDictionary[]] $Sections
    )
    $TypesRequired = New-ArrayList
    $Types = 'TableData', 'ListData', 'ChartData', 'SqlData', 'ExcelData', 'TextBasedData'
    foreach ($Section in $Sections) {
        $Keys = Get-ObjectKeys -Object $Section
        foreach ($Key in $Keys) {
            if ($Section.$Key.Use -eq $True) {
                foreach ($Type in $Types) {
                    #Write-Verbose "Get-TypesRequired - Section: $Key Type: $Type Value: $($Section.$Key.$Type)"
                    Add-ToArrayAdvanced -List $TypesRequired -Element $Section.$Key.$Type -SkipNull -RequireUnique -FullComparison
                }
            }
        }
    }
    Write-Verbose "Get-TypesRequired - FinalList: $($TypesRequired -join ', ')"
    return $TypesRequired
}