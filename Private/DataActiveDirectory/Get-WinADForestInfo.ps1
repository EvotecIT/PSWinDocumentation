function Get-WinADForestInfo {
    [CmdletBinding()]
    param(
        [PSCustomObject] $Forest
    )
    [ordered] @{
        'Name'                    = $Forest.Name
        'Root Domain'             = $Forest.RootDomain
        'Forest Functional Level' = $Forest.ForestMode
        'Domains Count'           = ($Forest.Domains).Count
        'Sites Count'             = ($Forest.Sites).Count
        'Domains'                 = ($Forest.Domains) -join ", "
        'Sites'                   = ($Forest.Sites) -join ", "
    }
}