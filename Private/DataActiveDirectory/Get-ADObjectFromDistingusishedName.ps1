function Get-ADObjectFromDistingusishedName {
    [CmdletBinding()]
    param (
        [string[]] $DistinguishedName,
        [Object[]] $ADCatalog,
        [string] $Type = '',
        [string] $Splitter # ', ' # Alternative for example [System.Environment]::NewLine
    )
    if ($DistinguishedName -eq $null) {
        return
    }
    $FoundObjects = foreach ($Catalog in $ADCatalog) {
        foreach ($Object in $DistinguishedName) {
            $ADObject = $Catalog | & { process { if ($_.DistinguishedName -eq $Object ) { $_ } } }  #| Where-Object { $_.DistinguishedName -eq $Object }
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