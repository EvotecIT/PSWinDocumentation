function Get-ADObjectFromDistingusishedName {
    param (
        [string[]] $ADObjectDistinguishedName,
        [Object[]] $ADCatalog,
        [string] $Type = 'SamAccountName',
        [string] $Splitter = ', ' # Alternative for example [System.Environment]::NewLine
    )
    $FoundObjects = @()

    foreach ($Catalog in $ADCatalog) {
        foreach ($Object in $ADObjectDistinguishedName) {
            $ADObject = $Catalog | Where { $_.DistinguishedName -eq $Object }
            if ($ADObject) {
                $FoundObjects += $ADObject.$Type
            }
        }
    }
    return ($FoundObjects | Sort-Object) -join $Splitter
}
