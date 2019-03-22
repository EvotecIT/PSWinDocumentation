function Get-WinADForestSubnets1 {
    param(
        [Array] $ForestSubnets
    )
    @(
        foreach ($Subnets in $ForestSubnets) {
            [PSCustomObject][ordered] @{
                'Name'        = $Subnets.Name
                'Description' = $Subnets.Description
                'Protected'   = $Subnets.ProtectedFromAccidentalDeletion
                'Modified'    = $Subnets.Modified
                'Created'     = $Subnets.Created
                'Deleted'     = $Subnets.Deleted
            }
        }
    )
}