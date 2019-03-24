function Get-WinADForest {
    [CmdletBinding()]
    param()
    $Time = Start-TimeLog
    Write-Verbose 'Getting forest information - Forest'

    Get-ADForest #| Select-Object -Property * -ExcludeProperty PropertyNames, AddedProperties, RemovedProperties, ModifiedProperties, PropertyCount

    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose "Getting forest information - Forest Time: $EndTime"
}