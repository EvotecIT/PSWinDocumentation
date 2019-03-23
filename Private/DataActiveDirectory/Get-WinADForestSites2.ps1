function Get-WinADForestSites2 {
    [CmdletBinding()]
    param(
        [Array] $ForestSites
    )
    @(
        foreach ($Sites in $ForestSites) {
            [PSCustomObject][ordered] @{
                'Name'                                = $Sites.Name
                'Topology Cleanup Enabled'            = $Sites.TopologyCleanupEnabled
                'Topology Detect Stale Enabled'       = $Sites.TopologyDetectStaleEnabled
                'Topology Minimum Hops Enabled'       = $Sites.TopologyMinimumHopsEnabled
                'Universal Group Caching Enabled'     = $Sites.UniversalGroupCachingEnabled
                'Universal Group Caching RefreshSite' = $Sites.UniversalGroupCachingRefreshSite
            }
        }
    )
}