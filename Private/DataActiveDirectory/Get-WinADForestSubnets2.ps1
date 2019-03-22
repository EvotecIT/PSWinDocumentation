function Get-WinADForestSubnets2 {
    param(
        [Array] $ForestSubnets
    )
    @(
        foreach ($Subnets in $ForestSubnets) {
            [PSCustomObject][ordered] @{
                'Name' = $Subnets.Name
                'Site' = $Subnets.Site
            }
        }
    )
}