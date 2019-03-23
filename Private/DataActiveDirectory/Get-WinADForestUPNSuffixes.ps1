function Get-WinADForestUPNSuffixes {
    param(
        [PSCustomObject] $Forest
    )
    @(
        [PSCustomObject] @{
            Name = $Forest.RootDomain
            Type = 'Primary / Default UPN'
        }
        foreach ($UPN in $Forest.UPNSuffixes) {
            [PSCustomObject] @{
                Name = $UPN
                Type = 'Secondary'
            }
        }
    )
}