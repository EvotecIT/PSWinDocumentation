function Get-WinADForestSites1 {
    [CmdletBinding()]
    param(
        [Array] $ForestSites
    )
    @(
        foreach ($Sites in $ForestSites) {
            [PSCustomObject][ordered] @{
                'Name'        = $Sites.Name
                'Description' = $Sites.Description
                #'sD Rights Effective'                = $Sites.sDRightsEffective
                'Protected'   = $Sites.ProtectedFromAccidentalDeletion
                'Modified'    = $Sites.Modified
                'Created'     = $Sites.Created
                'Deleted'     = $Sites.Deleted
            }
        }
    )
}