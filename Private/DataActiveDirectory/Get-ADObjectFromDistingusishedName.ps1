function Get-ADObjectFromDistingusishedName {
    [CmdletBinding()]
    param (
        [string[]] $DistinguishedName,
        [Object[]] $ADCatalog,
        [string] $Type = '',
        [string] $Splitter # ', ' # Alternative for example [System.Environment]::NewLine
    )
    $FoundObjects = foreach ($Catalog in $ADCatalog) {
        foreach ($Object in $DistinguishedName) {
            $ADObject = $Catalog | Where-Object { $_.DistinguishedName -eq $Object }
            if ($ADObject) {
                if ($Type -eq '') {
                    #Write-Verbose 'Get-ADObjectFromDistingusishedName - Whole object'
                    $ADObject
                } else {
                    #Write-Verbose 'Get-ADObjectFromDistingusishedName - Part of object'
                    $ADObject.$Type
                }
            }
        }
    }
    if ($Splitter) {
        return ($FoundObjects | Sort-Object) -join $Splitter
    } else {
        return $FoundObjects | Sort-Object
    }
}
