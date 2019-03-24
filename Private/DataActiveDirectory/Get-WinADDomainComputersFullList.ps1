function Get-WinADDomainComputersFullList {
    [cmdletbinding()]
    param(
        [string] $Domain
    )
    Write-Verbose "Getting domain information - $Domain DomainComputersFullList"
    $TimeUsers = Start-TimeLog
    [string[]] $Properties = '*'
    [string[]] $ExcludeProperty = '*Certificate', 'PropertyNames', '*Properties', 'PropertyCount', 'Certificates', 'nTSecurityDescriptor'

    Get-ADComputer -Server $Domain -Filter * -ResultPageSize 500000 -Properties $Properties | Select-Object -Property $Properties -ExcludeProperty $ExcludeProperty

    $EndUsers = Stop-TimeLog -Time $TimeUsers -Option OneLiner
    Write-Verbose "Getting domain information - $Domain DomainComputersFullList Time: $EndUsers"
}