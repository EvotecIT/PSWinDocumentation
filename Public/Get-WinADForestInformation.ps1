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
    $Data.Forest = $(Get-ADForest)
    Write-Verbose 'Getting forest information - RootDSE'
    $Data.RootDSE = $(Get-ADRootDSE -Properties *)
    Write-Verbose 'Getting forest information - ForestName/ForestNameDN'
    $Data.ForestName = $Data.Forest.Name
    $Data.ForestNameDN = $Data.RootDSE.defaultNamingContext
    $Data.Domains = $Data.Forest.Domains

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSites, [ActiveDirectory]::ForestSites1, [ActiveDirectory]::ForestSites2)) {
        Write-Verbose 'Getting forest information - Forest Sites'
        $Data.ForestSites = $(Get-ADReplicationSite -Filter * -Properties * )
        $Data.ForestSites1 = Invoke-Command -ScriptBlock {
            $ReturnData = @()
            foreach ($Sites in $Data.ForestSites) {
                $ReturnData += [ordered] @{
                    'Name'        = $Sites.Name
                    'Description' = $Sites.Description
                    #'sD Rights Effective'                = $Sites.sDRightsEffective
                    'Protected'   = $Sites.ProtectedFromAccidentalDeletion
                    'Modified'    = $Sites.Modified
                    'Created'     = $Sites.Created
                    'Deleted'     = $Sites.Deleted
                }
            }
            return Format-TransposeTable $ReturnData
        }
        $Data.ForestSites2 = Invoke-Command -ScriptBlock {
            $ReturnData = @()
            foreach ($Sites in $Data.ForestSites) {
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
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSubnet , [ActiveDirectory]::ForestSubnets1, [ActiveDirectory]::ForestSubnets2)) {
        Write-Verbose 'Getting forest information - Forest Subnets'
        $Data.ForestSubnets = $(Get-ADReplicationSubnet -Filter * -Properties * | `
                Select-Object  Name, DisplayName, Description, Site, ProtectedFromAccidentalDeletion, Created, Modified, Deleted )
        $Data.ForestSubnets1 = Invoke-Command -ScriptBlock {
            $ReturnData = @()
            foreach ($Subnets in $Data.ForestSubnets) {
                $ReturnData += [ordered] @{
                    'Name'        = $Subnets.Name
                    'Description' = $Subnets.Description
                    'Protected'   = $Subnets.ProtectedFromAccidentalDeletion
                    'Modified'    = $Subnets.Modified
                    'Created'     = $Subnets.Created
                    'Deleted'     = $Subnets.Deleted
                }
            }
            return Format-TransposeTable $ReturnData
        }
        $Data.ForestSubnets2 = Invoke-Command -ScriptBlock {
            $ReturnData = @()
            foreach ($Subnets in $Data.ForestSubnets) {
                $ReturnData += [ordered] @{
                    'Name' = $Subnets.Name
                    'Site' = $Subnets.Site
                }
            }
            return Format-TransposeTable $ReturnData
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
            $UPNSuffixList = @()
            $UPNSuffixList += $Data.Forest.RootDomain + ' (Primary / Default UPN)'
            if ($Data.Forest.UPNSuffixes) {
                $UPNSuffixList += $Data.Forest.UPNSuffixes
            }
            return $UPNSuffixList
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
            $FSMO = [ordered] @{
                'Domain Naming Master' = $Data.Forest.DomainNamingMaster
                'Schema Master'        = $Data.Forest.SchemaMaster
            }
            return $FSMO
        }
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestOptionalFeatures)) {
        Write-Verbose 'Getting forest information - Forest Optional Features'
        $Data.ForestOptionalFeatures = Invoke-Command -ScriptBlock {
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
    }
    ### Generate Data from Domains
    $Data.FoundDomains = [ordered]@{}
    $DomainData = @()
    foreach ($Domain in $Data.Domains) {
        $Data.FoundDomains.$Domain = Get-WinADDomainInformation -Domain $Domain -TypesRequired $TypesRequired -PathToPasswords $PathToPasswords -PathToPasswordsHashes $PathToPasswordsHashes
    }
    return $Data
}
