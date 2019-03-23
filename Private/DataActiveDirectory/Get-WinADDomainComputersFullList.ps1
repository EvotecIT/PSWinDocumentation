function Get-WinADDomainComputersFullList {
    [cmdletbinding()]
    param(
        [string] $Domain
    )
    [string[]] $Properties = '*'
    [string[]] $ExcludeProperty = '*Certificate', 'PropertyNames', '*Properties', 'PropertyCount', 'Certificates', 'nTSecurityDescriptor'
    return Get-ADComputer -Server $Domain -Filter * -ResultPageSize 500000 -Properties $Properties | Select-Object -Property $Properties -ExcludeProperty $ExcludeProperty
}