function Get-WinADRootDSE {
    [CmdletBinding()]
    param()
    $Time = Start-TimeLog
    Write-Verbose 'Getting forest information - RootDSE'

    Get-ADRootDSE -Properties * #| Select-Object -Property * -ExcludeProperty PropertyNames, AddedProperties, RemovedProperties, ModifiedProperties, PropertyCount

    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose "Getting forest information - RootDSE Time: $EndTime"
}