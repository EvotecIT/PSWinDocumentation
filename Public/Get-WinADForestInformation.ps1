function Get-WinADForestInformation {
    [CmdletBinding()]
    param (
        [Object] $TypesRequired,
        [switch] $RequireTypes,
        [string] $PathToPasswords,
        [string] $PathToPasswordsHashes
    )
    if ($TypesRequired -eq $null) {
        Write-Verbose 'Get-WinADForestInformation - TypesRequired is null. Getting all.'
        $TypesRequired = Get-Types -Types ([ActiveDirectory])
    } # Gets all types

    $Data = [ordered] @{}
    Write-Verbose 'Getting forest information - Forest'
    $Data.Forest = Get-WinForest
    Write-Verbose 'Getting forest information - RootDSE'
    $Data.RootDSE = Get-WinADRootDSE
    Write-Verbose 'Getting forest information - ForestName/ForestNameDN'
    $Data.ForestName = $Data.Forest.Name
    $Data.ForestNameDN = $Data.RootDSE.defaultNamingContext
    $Data.Domains = $Data.Forest.Domains

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSites, [ActiveDirectory]::ForestSites1, [ActiveDirectory]::ForestSites2)) {
        Write-Verbose 'Getting forest information - Forest Sites'
        $Data.ForestSites = Get-WinADForestSites
        $Data.ForestSites1 = Invoke-Command -ScriptBlock {
            @(
                foreach ($Sites in $Data.ForestSites) {
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
        $Data.ForestSites2 = Invoke-Command -ScriptBlock {
            @(
                foreach ($Sites in $Data.ForestSites) {
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
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSubnet , [ActiveDirectory]::ForestSubnets1, [ActiveDirectory]::ForestSubnets2)) {
        Write-Verbose 'Getting forest information - Forest Subnets'
        $Data.ForestSubnets = $(Get-ADReplicationSubnet -Filter * -Properties * | `
                Select-Object  Name, DisplayName, Description, Site, ProtectedFromAccidentalDeletion, Created, Modified, Deleted )
        $Data.ForestSubnets1 = Invoke-Command -ScriptBlock {
            @(
                foreach ($Subnets in $Data.ForestSubnets) {
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
        $Data.ForestSubnets2 = Invoke-Command -ScriptBlock {
            @(
                foreach ($Subnets in $Data.ForestSubnets) {
                    [PSCustomObject][ordered] @{
                        'Name' = $Subnets.Name
                        'Site' = $Subnets.Site
                    }
                }
            )
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSiteLinks)) {
        Write-Verbose 'Getting forest information - Forest SiteLinks'
        $Data.ForestSiteLinks = $(
            Get-ADReplicationSiteLink -Filter * -Properties `
                Name, Cost, ReplicationFrequencyInMinutes, replInterval, ReplicationSchedule, Created, Modified, Deleted, IsDeleted, ProtectedFromAccidentalDeletion | `
                Select-Object Name, Cost, ReplicationFrequencyInMinutes, ReplInterval, Modified
        )
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestInformation)) {
        Write-Verbose 'Getting forest information - Forest Information'
        $Data.ForestInformation = [ordered] @{
            'Name'                    = $Data.Forest.Name
            'Root Domain'             = $Data.Forest.RootDomain
            'Forest Functional Level' = $Data.Forest.ForestMode
            'Domains Count'           = ($Data.Forest.Domains).Count
            'Sites Count'             = ($Data.Forest.Sites).Count
            'Domains'                 = ($Data.Forest.Domains) -join ", "
            'Sites'                   = ($Data.Forest.Sites) -join ", "
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestUPNSuffixes)) {
        Write-Verbose 'Getting forest information - Forest UPNSuffixes'
        $Data.ForestUPNSuffixes = Invoke-Command -ScriptBlock {
            @(
                $Data.Forest.RootDomain + ' (Primary / Default UPN)'
                if ($Data.Forest.UPNSuffixes) {
                    $Data.Forest.UPNSuffixes
                }
            )           
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestGlobalCatalogs)) {
        Write-Verbose 'Getting forest information - Forest GlobalCatalogs'
        $Data.ForestGlobalCatalogs = $Data.Forest.GlobalCatalogs
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSPNSuffixes)) {
        Write-Verbose 'Getting forest information - Forest SPNSuffixes'
        $Data.ForestSPNSuffixes = $Data.Forest.SPNSuffixes
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestFSMO)) {
        Write-Verbose 'Getting forest information - Forest FSMO'
        $Data.ForestFSMO = Invoke-Command -ScriptBlock {
            [ordered] @{
                'Domain Naming Master' = $Data.Forest.DomainNamingMaster
                'Schema Master'        = $Data.Forest.SchemaMaster
            }            
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestOptionalFeatures)) {
        Write-Verbose 'Getting forest information - Forest Optional Features'
        $Data.ForestOptionalFeatures = Get-WinADForestOptionalFeatures
    }
    ### Generate Data from Domains
    $Data.FoundDomains = [ordered]@{}
    #$DomainData = @()
    foreach ($Domain in $Data.Domains) {
        $Data.FoundDomains.$Domain = Get-WinADDomainInformation -Domain $Domain -TypesRequired $TypesRequired -PathToPasswords $PathToPasswords -PathToPasswordsHashes $PathToPasswordsHashes
    }
    return $Data
}
