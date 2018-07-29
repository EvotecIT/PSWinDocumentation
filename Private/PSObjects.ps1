function Get-ObjectCount {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Object]$Object
    )
    return $($Object | Measure-Object).Count
}
function Convert-ToTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Object]$Object
    )
    return $Object.ForEach( {[PSCustomObject]$_})
}