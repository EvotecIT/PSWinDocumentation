function Get-WinADDomainGroupsFullList {
    [CmdletBinding()]
    param(
        [string] $Domain
    )
    [string[]] $Properties = '*'
    [string[]] $ExcludeProperty = '*Certificate', 'PropertyNames', '*Properties', 'PropertyCount', 'Certificates', 'nTSecurityDescriptor'
    return Get-ADGroup -Server $Domain -Filter * -ResultPageSize 500000 -Properties $Properties | Select-Object -Property $Properties -ExcludeProperty $ExcludeProperty
}