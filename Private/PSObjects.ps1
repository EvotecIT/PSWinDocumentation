function Get-ObjectCount {
    [CmdletBinding()]
    param(
        $Object
    )
    return $($Object | Measure-Object).Count
}