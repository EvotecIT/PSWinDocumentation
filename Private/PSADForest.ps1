function Get-WinADForestInformation {
    [CmdletBinding()]
    $Data = @{}
    $ForestInformation = $(Get-ADForest)
    $Data.Forest = $ForestInformation
    $Data.RootDSE = $(Get-ADRootDSE -Properties *)
    $Data.Sites = $(Get-ADReplicationSite -Filter * -Properties * )
    <#
    $Data.Sites1 = $(
        $Data.Sites | Select-Object Name, Description, Modified sDRightsEffective, ProtectedFromAccidentalDeletion, Created, Modified, Deleted
    )
    #>
    $Data.Sites1 = Invoke-Command -ScriptBlock {
        $ReturnData = @()
        foreach ($Sites in $Data.Sites) {
            $ReturnData += [ordered] @{
                'Name'                               = $Sites.Name
                'Description'                        = $Sites.Description
                'sD Rights Effective'                = $Sites.sDRightsEffective
                'Protected From Accidental Deletion' = $Sites.ProtectedFromAccidentalDeletion
                'Modified'                           = $Sites.Modified
                'Created'                            = $Sites.Created
                'Deleted'                            = $Sites.Deleted
            }
        }
        return Format-TransposeTable $ReturnData
    }
    <#
    $Data.Sites2 = $(
        $Data.Sites | Select-Object Name, TopologyCleanupEnabled, TopologyDetectStaleEnabled, TopologyMinimumHopsEnabled, UniversalGroupCachingEnabled, UniversalGroupCachingRefreshSite
    )
    $Data.Sites3 = Invoke-Command -ScriptBlock {
        $ReturnData = [ordered] @{
            'Name'                             = $Data.Sites.Name
            'TopologyCleanupEnabled'           = $Data.Sites.TopologyCleanupEnabled
            'TopologyDetectStaleEnabled'       = $Data.Sites.TopologyDetectStaleEnabled
            'TopologyMinimumHopsEnabled'       = $Data.Sites.TopologyMinimumHopsEnabled
            'UniversalGroupCachingEnabled'     = $Data.Sites.UniversalGroupCachingEnabled
            'UniversalGroupCachingRefreshSite' = $Data.Sites.UniversalGroupCachingRefreshSite
        }
        return $ReturnData | Convert-ToTable
    }
    #>
    $Data.Sites2 = Invoke-Command -ScriptBlock {
        $ReturnData = @()
        foreach ($Sites in $Data.Sites) {
            $ReturnData += [ordered] @{
                'Name'                                = $Sites.Name
                'Topology Cleanup Enabled'            = $Sites.TopologyCleanupEnabled
                'Topology Detect Stale Enabled'       = $Sites.TopologyDetectStaleEnabled
                'Topology Minimum Hops Enabled'       = $Sites.TopologyMinimumHopsEnabled
                'Universal Group Caching Enabled'     = $Sites.UniversalGroupCachingEnabled
                'Universal Group Caching RefreshSite' = $Sites.UniversalGroupCachingRefreshSite
            }
        }
        return Format-TransposeTable $ReturnData
    }

    $Data.Subnets = $(Get-ADReplicationSubnet -Filter * -Properties * | `
            Select-Object  Name, DisplayName, Description, Site, ProtectedFromAccidentalDeletion, Created, Modified, Deleted )
    $Data.Subnets1 = Invoke-Command -ScriptBlock {
        $ReturnData = @()
        foreach ($Subnets in $Data.Subnets) {
            $ReturnData += [ordered] @{
                'Name'                               = $Subnets.Name
                'Description'                        = $Subnets.Description
                'Protected From Accidental Deletion' = $Subnets.ProtectedFromAccidentalDeletion
                'Modified'                           = $Subnets.Modified
                'Created'                            = $Subnets.Created
                'Deleted'                            = $Subnets.Deleted
            }
        }
        return Format-TransposeTable $ReturnData
    }
    $Data.Subnets2 = Invoke-Command -ScriptBlock {
        $ReturnData = @()
        foreach ($Subnets in $Data.Subnets) {
            $ReturnData += [ordered] @{
                'Name' = $Subnets.Name
                'Site' = $Subnets.Site
            }
        }
        return Format-TransposeTable $ReturnData
    }




    $Data.SiteLinks = $(
        Get-ADReplicationSiteLink -Filter * -Properties `
            Name, Cost, ReplicationFrequencyInMinutes, replInterval, ReplicationSchedule, Created, Modified, Deleted, IsDeleted, ProtectedFromAccidentalDeletion | `
            Select-Object Name, Cost, ReplicationFrequencyInMinutes, ReplInterval, Modified
    )

    $Data.ForestName = $ForestInformation.Name
    $Data.ForestNameDN = $Data.RootDSE.defaultNamingContext
    $Data.Domains = $ForestInformation.Domains
    $Data.ForestInformation = [ordered] @{
        'Name'                    = $ForestInformation.Name
        'Root Domain'             = $ForestInformation.RootDomain
        'Forest Functional Level' = $ForestInformation.ForestMode
        'Domains Count'           = ($ForestInformation.Domains).Count
        'Sites Count'             = ($ForestInformation.Sites).Count
        'Domains'                 = ($ForestInformation.Domains) -join ", "
        'Sites'                   = ($ForestInformation.Sites) -join ", "
    }
    $Data.UPNSuffixes = Invoke-Command -ScriptBlock {
        $UPNSuffixList = @()
        $UPNSuffixList += $ForestInformation.RootDomain + ' (Primary / Default UPN)'
        $UPNSuffixList += $ForestInformation.UPNSuffixes
        return $UPNSuffixList
    }
    $Data.GlobalCatalogs = $ForestInformation.GlobalCatalogs
    $Data.SPNSuffixes = $ForestInformation.SPNSuffixes
    $Data.FSMO = Invoke-Command -ScriptBlock {
        $FSMO = [ordered] @{
            'Domain Naming Master' = $ForestInformation.DomainNamingMaster
            'Schema Master'        = $ForestInformation.SchemaMaster
        }
        return $FSMO
    }
    $Data.OptionalFeatures = Invoke-Command -ScriptBlock {
        $OptionalFeatures = $(Get-ADOptionalFeature -Filter * )
        $Optional = [ordered]@{
            'Recycle Bin Enabled'                          = ''
            'Privileged Access Management Feature Enabled' = ''
        }
        ### Fix Optional Features
        foreach ($Feature in $OptionalFeatures) {
            if ($Feature.Name -eq 'Recycle Bin Feature') {
                if ("$($Feature.EnabledScopes)" -eq '') {
                    $Optional.'Recycle Bin Enabled' = $False
                } else {
                    $Optional.'Recycle Bin Enabled' = $True
                }
            }
            if ($Feature.Name -eq 'Privileged Access Management Feature') {
                if ("$($Feature.EnabledScopes)" -eq '') {
                    $Optional.'Privileged Access Management Feature Enabled' = $False
                } else {
                    $Optional.'Privileged Access Management Feature Enabled' = $True
                }
            }
        }
        return $Optional
        ### Fix optional features
    }
    return $Data
}