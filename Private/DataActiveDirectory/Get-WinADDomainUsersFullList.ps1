function Get-WinADDomainUsersFullList {
    [CmdletBinding()]
    param(
        [string] $Domain
    )
    Write-Verbose "Getting domain information - $Domain DomainUsersFullList"
    $TimeUsers = Start-TimeLog
    [string[]] $Properties = '*' #, "msDS-UserPasswordExpiryTimeComputed"
    [string[]] $ExcludeProperty = '*Certificate', 'PropertyNames', '*Properties', 'PropertyCount', 'Certificates', 'nTSecurityDescriptor'

    Get-ADUser -Server $Domain -ResultPageSize 500000 -Filter * -Properties $Properties | Select-Object -Property $Properties -ExcludeProperty $ExcludeProperty

    $EndUsers = Stop-TimeLog -Time $TimeUsers -Option OneLiner
    Write-Verbose "Getting domain information - $Domain DomainUsersFullList Time: $EndUsers"
}