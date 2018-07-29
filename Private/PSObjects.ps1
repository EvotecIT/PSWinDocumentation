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
function Convert-KeyToKeyValue {
    param (
        $Object
    )
    $NewHash = [ordered] @{}
    foreach ($O in $Object.Keys) {
        $KeyName = "$O ($($Object.$O))"
        $KeyValue = $Object.$O
        $NewHash.$KeyName = $KeyValue
    }
    return $NewHash
}