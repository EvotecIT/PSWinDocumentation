function Get-ObjectCount {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Object]$Object
    )
    return $($Object | Measure-Object).Count
}
function Convert-KeyToKeyValue {
    [CmdletBinding()]
    param (
        [object] $Object
    )
    $NewHash = [ordered] @{}
    foreach ($O in $Object.Keys) {
        $KeyName = "$O ($($Object.$O))"
        $KeyValue = $Object.$O
        $NewHash.$KeyName = $KeyValue
    }
    return $NewHash
}
function Get-ObjectKeys {
    param(
        [object] $Object,
        [string] $Ignore
    )
    $Data = $Object.Keys | Where { $_ -notcontains $Ignore }
    return $Data
}

## This methods converts 2 Arrays into 1 Array
## Administrators  + 0 = Administrators (0)
function Convert-TwoArraysIntoOne {
    [CmdletBinding()]
    param (
        $Object,
        $ObjectToAdd
    )

    $Value = @()
    for ($i = 0; $i -lt $Object.Count; $i++) {
        $Value += "$($Object[$i]) ($($ObjectToAdd[$i]))"
    }
    return $Value
}
