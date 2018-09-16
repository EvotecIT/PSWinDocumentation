Function Get-Types {
    [CmdletBinding()]
    param (
        [Object] $Types
    )
    $TypesRequired = @()
    foreach ($Type in $Types) {
        #Write-Verbose "Type: $Type"
        $TypesRequired += $Type.GetEnumValues()
    }
    return $TypesRequired
}