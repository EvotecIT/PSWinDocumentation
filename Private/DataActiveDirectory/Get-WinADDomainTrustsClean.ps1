function Get-WinADDomainTrustsClean {
    [CmdletBinding()]
    param(
        [string] $Domain,
        [Array] $TypesRequired
    )
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainTrustsClean, [ActiveDirectory]::DomainTrusts)) {
        Write-Verbose "Getting domain information - $Domain DomainTrustsClean"
        $Time = Start-TimeLog

        Get-ADTrust -Server $Domain -Filter * -Properties * -ErrorAction SilentlyContinue

        $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
        Write-Verbose "Getting domain information - $Domain DomainTrustsClean Time: $EndTime"
    }
}