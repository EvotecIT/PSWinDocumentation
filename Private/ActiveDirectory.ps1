function Get-WinADForestInformation {
    [CmdletBinding()]
    param (
        $TypesRequired
    )
    $Data = [ordered] @{}
    $Data.Forest = $(Get-ADForest)
    $Data.RootDSE = $(Get-ADRootDSE -Properties *)
    $Data.Sites = $(Get-ADReplicationSite -Filter * -Properties * )
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
    $Data.ForestName = $Data.Forest.Name
    $Data.ForestNameDN = $Data.RootDSE.defaultNamingContext
    $Data.Domains = $Data.Forest.Domains
    $Data.ForestInformation = [ordered] @{
        'Name'                    = $Data.Forestn.Name
        'Root Domain'             = $Data.Forest.RootDomain
        'Forest Functional Level' = $Data.Forest.ForestMode
        'Domains Count'           = ($Data.Forest.Domains).Count
        'Sites Count'             = ($Data.Forest.Sites).Count
        'Domains'                 = ($Data.Forest.Domains) -join ", "
        'Sites'                   = ($Data.Forest.Sites) -join ", "
    }
    $Data.UPNSuffixes = Invoke-Command -ScriptBlock {
        $UPNSuffixList = @()
        $UPNSuffixList += $Data.Forest.RootDomain + ' (Primary / Default UPN)'
        $UPNSuffixList += $Data.Forest.UPNSuffixes
        return $UPNSuffixList
    }
    $Data.GlobalCatalogs = $Data.Forest.GlobalCatalogs
    $Data.SPNSuffixes = $Data.Forest.SPNSuffixes
    $Data.FSMO = Invoke-Command -ScriptBlock {
        $FSMO = [ordered] @{
            'Domain Naming Master' = $Data.Forest.DomainNamingMaster
            'Schema Master'        = $Data.Forest.SchemaMaster
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
    ### Generate Data from Domains
    $Data.FoundDomains = [ordered]@{}
    $DomainData = @()
    foreach ($Domain in $Data.Domains) {
        $Data.FoundDomains.$Domain = Get-WinADDomainInformation -Domain $Domain -TypesRequired $TypesRequired
    }
    return $Data
}

function Get-WinADDomainInformation {
    [CmdletBinding()]
    param (
        [string] $Domain,
        [Object] $TypesRequired
    )
    if ($TypesRequired -eq $null) {
        Write-Verbose ' Get-WinADDomainInformation - TypesRequired is null. Getting all.'
        $TypesRequired = Get-Types
    } # Gets all types
    $Data = [ordered] @{}
    $Data.RootDSE = $(Get-ADRootDSE -Server $Domain)
    $Data.DomainInformation = $(Get-ADDomain -Server $Domain)
    if ($TypesRequired -contains [Domain]::AuthenticationPolicies) {
        $Data.AuthenticationPolicies = $(Get-ADAuthenticationPolicy -Server $Domain -LDAPFilter '(name=AuthenticationPolicy*)')
    }
    if ($TypesRequired -contains [Domain]::AuthenticationPolicySilos) {
        $Data.AuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Server $Domain -Filter 'Name -like "*AuthenticationPolicySilo*"')
    }
    if ($TypesRequired -contains [Domain]::CentralAccessPolicies) {
        $Data.CentralAccessPolicies = $(Get-ADCentralAccessPolicy -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [Domain]::CentralAccessRules) {
        $Data.CentralAccessRules = $(Get-ADCentralAccessRule -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [Domain]::ClaimTransformPolicies) {
        $Data.ClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [Domain]::ClaimTypes) {
        $Data.ClaimTypes = $(Get-ADClaimType -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [Domain]::LDAPDNS) {
        $Data.LDAPDNS = $(Resolve-DnsName -Name "_ldap._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    }
    if ($TypesRequired -contains [Domain]::KerberosDNS) {
        $Data.KerberosDNS = $(Resolve-DnsName -Name "_kerberos._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    }
    if ($TypesRequired -contains [Domain]::FSMO -or $TypesRequired -contains [Domain]::DomainTrusts) {
        # required for multiple use cases FSMO/DomainTrusts
        $Data.FSMO = [ordered] @{
            'PDC Emulator'          = $Data.DomainInformation.PDCEmulator
            'RID Master'            = $Data.DomainInformation.RIDMaster
            'Infrastructure Master' = $Data.DomainInformation.InfrastructureMaster
        }
    }
    if ($TypesRequired -contains [Domain]::DomainTrusts) {
        ## requires both DomainTrusts and FSMO.
        $Data.DomainTrustsClean = (Get-ADTrust -Server $Domain -Filter * -Properties *)
        $Data.DomainTrusts = Invoke-Command -ScriptBlock {
            $DomainPDC = $Data.FSMO.'PDC Emulator'
            $Trust = $Data.DomainTrustsClean
            $TrustWMI = Get-CimInstance -ClassName Microsoft_DomainTrustStatus -Namespace root\MicrosoftActiveDirectory -ComputerName $DomainPDC -ErrorAction SilentlyContinue | Select-Object TrustIsOK, TrustStatus, TrustStatusString, PSComputerName, TrustedDCName

            $ReturnData = [ordered] @{
                'Trust Source'               = $Domain
                'Trust Target'               = $Trust.Target
                'Trust Direction'            = $Trust.Direction
                'Trust Attributes'           = Set-TrustAttributes -Value $Trust.TrustAttributes
                #'Trust OK'                   = $TrustWMI.TrustIsOK
                #'Trust Status'               = $TrustWMI.TrustStatus
                'Trust Status'               = $TrustWMI.TrustStatusString
                'Forest Transitive'          = $Trust.ForestTransitive
                'Selective Authentication'   = $Trust.SelectiveAuthentication
                'SID Filtering Forest Aware' = $Trust.SIDFilteringForestAware
                'SID Filtering Quarantined'  = $Trust.SIDFilteringQuarantined
                'Disallow Transivity'        = $Trust.DisallowTransivity
                'Intra Forest'               = $Trust.IntraForest
                'Is Tree Parent'             = $Trust.IsTreeParent
                'Is Tree Root'               = $Trust.IsTreeRoot
                'TGTDelegation'              = $Trust.TGTDelegation
                'TrustedPolicy'              = $Trust.TrustedPolicy
                'TrustingPolicy'             = $Trust.TrustingPolicy
                'TrustType'                  = $Trust.TrustType
                'UplevelOnly'                = $Trust.UplevelOnly
                'UsesAESKeys'                = $Trust.UsesAESKeys
                'UsesRC4Encryption'          = $Trust.UsesRC4Encryption
                'Trust Source DC'            = $TrustWMI.PSComputerName
                'Trust Target DC'            = $TrustWMI.TrustedDCName.Replace('\\', '')
                'Trust Source DN'            = $Trust.Source
                'ObjectGUID'                 = $Trust.ObjectGUID
                'Created'                    = $Trust.Created
                'Modified'                   = $Trust.Modified
                'Deleted'                    = $Trust.Deleted
                'SID'                        = $Trust.securityIdentifier
            }
            return Format-TransposeTable $ReturnData
        }
    }
    if ($TypesRequired -contains [Domain]::GroupPolicies) {
        $Data.GroupPoliciesClean = $(Get-GPO -Domain $Domain -All)
        $Data.GroupPolicies = Invoke-Command -ScriptBlock {
            $GroupPolicies = @()
            foreach ($gpo in $Data.GroupPoliciesClean) {
                $GroupPolicy = [ordered] @{
                    'Display Name'      = $gpo.DisplayName
                    'Gpo Status'        = $gpo.GPOStatus
                    'Creation Time'     = $gpo.CreationTime
                    'Modification Time' = $gpo.ModificationTime
                    'Description'       = $gpo.Description
                    'Wmi Filter'        = $gpo.WmiFilter
                }
                $GroupPolicies += $GroupPolicy
            }
            return Format-TransposeTable $GroupPolicies
        }
    }
    if ($TypesRequired -contains [Domain]::GroupPoliciesDetails) {
        $Data.GroupPoliciesDetails = Format-TransposeTable (Get-GPOInfo -DomainName $Domain)
    }
    if ($TypesRequired -contains [Domain]::DefaultPasswordPolicy) {
        $Data.DefaultPassWordPoLicy = Invoke-Command -ScriptBlock {
            $DefaultPasswordPolicy = $(Get-ADDefaultDomainPasswordPolicy -Server $Domain)
            $Data = [ordered] @{
                'Complexity Enabled'            = $DefaultPasswordPolicy.ComplexityEnabled
                'Lockout Duration'              = $DefaultPasswordPolicy.LockoutDuration
                'Lockout Observation Window'    = $DefaultPasswordPolicy.LockoutObservationWindow
                'Lockout Threshold'             = $DefaultPasswordPolicy.LockoutThreshold
                'Max Password Age'              = $DefaultPasswordPolicy.MaxPasswordAge
                'Min Password Length'           = $DefaultPasswordPolicy.MinPasswordLength
                'Min Password Age'              = $DefaultPasswordPolicy.MinPasswordAge
                'Password History Count'        = $DefaultPasswordPolicy.PasswordHistoryCount
                'Reversible Encryption Enabled' = $DefaultPasswordPolicy.ReversibleEncryptionEnabled
                'Distinguished Name'            = $DefaultPasswordPolicy.DistinguishedName
            }
            return $Data
        }
    }
    if ($TypesRequired -contains [Domain]::PriviligedGroupMembers) {
        $Data.PriviligedGroupMembers = Get-PrivilegedGroupsMembers -Domain $Data.DomainInformation.DNSRoot -DomainSID $Data.DomainInformation.DomainSid
    }
    if ($TypesRequired -contains [Domain]::OrganizationalUnits) {
        $Data.OrganizationalUnitsClean = $(Get-ADOrganizationalUnit -Server $Domain -Properties * -Filter * )
        $Data.OrganizationalUnits = Invoke-Command -ScriptBlock {
            return $Data.OrganizationalUnitsClean | Select-Object CanonicalName, Created | Sort-Object CanonicalName
        }
    }
    if ($TypesRequired -contains [Domain]::DomainAdministrators) {
        $Data.DomainAdministratorsClean = $( Get-ADGroup -Server $Domain -Identity $('{0}-512' -f (Get-ADDomain -Server $Domain).domainSID) | Get-ADGroupMember -Server $Domain -Recursive | Get-ADUser -Server $Domain)
        $Data.DomainAdministrators = $Data.DomainAdministratorsClean | Select-Object Name, SamAccountName, UserPrincipalName, Enabled
    }
    if ($TypesRequired -contains [Domain]::Users -or $TypesRequired -contains [Domain]::UsersCount) {
        Write-Verbose 'Get-WinDomainInformation - Getting All Users'
        $Data.Users = Invoke-Command -ScriptBlock {
            param(
                $Domain
            )
            function Find-AllUsers {
                param (
                    $Domain
                )
                $users = Get-ADUser -Server $Domain -ResultPageSize 5000000 -filter * -Properties Name, Manager, DisplayName, GivenName, Surname, SamAccountName, EmailAddress, msDS-UserPasswordExpiryTimeComputed, PasswordExpired, PasswordLastSet, PasswordNotRequired, PasswordNeverExpires
                $users = $users | Select-Object Name, UserPrincipalName, SamAccountName, DisplayName, GivenName, Surname, EmailAddress, PasswordExpired, PasswordLastSet, PasswordNotRequired, PasswordNeverExpires, Enabled,
                @{Name = "Manager"; Expression = { (Get-ADUser -Server $Domain $_.Manager).Name }},
                @{Name = "ManagerEmail"; Expression = { (Get-ADUser -Server $Domain -Properties Mail $_.Manager).Mail  }},
                @{Name = "DateExpiry"; Expression = { ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")) }},
                @{Name = "DaysToExpire"; Expression = { (NEW-TIMESPAN -Start (GET-DATE) -End ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))).Days }}
                return $users
            }
            $Users = Find-AllUsers -Domain $Domain
            return [ordered] @{
                Users                          = $Users
                UsersAll                       = $Users | Where { $_.PasswordNotRequired -eq $False } | Select Name, SamAccountName, UserPrincipalName, Enabled
                UsersSystemAccounts            = $Users | Where { $_.PasswordNotRequired -eq $true } | Select Name, SamAccountName, UserPrincipalName, Enabled
                UsersNeverExpiring             = $Users | Where { $_.PasswordNeverExpires -eq $true -and $_.Enabled -eq $true -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
                UsersNeverExpiringInclDisabled = $Users | Where { $_.PasswordNeverExpires -eq $true -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
                UsersExpiredInclDisabled       = $Users | Where { $_.PasswordNeverExpires -eq $false -and $_.DaysToExpire -le 0 -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
                UsersExpiredExclDisabled       = $Users | Where { $_.PasswordNeverExpires -eq $false -and $_.DaysToExpire -le 0 -and $_.Enabled -eq $true -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
            }
        } -ArgumentList $Domain
    }
    if ($TypesRequired -contains [Domain]::UsersCount) {
        Write-Verbose 'Get-WinDomainInformation - Getting All Users Count'
        $Data.UsersCount = [ordered] @{
            'Users Count Incl. System'            = Get-ObjectCount -Object $Data.Users.Users
            'Users Count'                         = Get-ObjectCount -Object $Data.Users.UsersAll
            'Users Expired'                       = Get-ObjectCount -Object $Data.Users.UsersExpiredExclDisabled
            'Users Expired Incl. Disabled'        = Get-ObjectCount -Object $Data.Users.UsersExpiredInclDisabled
            'Users Never Expiring'                = Get-ObjectCount -Object $Data.Users.UsersNeverExpiring
            'Users Never Expiring Incl. Disabled' = Get-ObjectCount -Object $Data.Users.UsersNeverExpiringInclDisabled
            'Users System Accounts'               = Get-ObjectCount -Object $Data.Users.UsersSystemAccounts
        }
    }
    if ($TypesRequired -contains [Domain]::DomainControllers) {
        $Data.DomainControllersClean = $(Get-ADDomainController -Server $Domain -Filter * )
        $Data.DomainControllers = Invoke-Command -ScriptBlock {
            $DCs = @()
            foreach ($DC in $Data.DomainControllersClean) {
                $DCs += [ordered] @{
                    'Name'               = $DC.Name
                    'Host Name'          = $DC.HostName
                    'Operating System'   = $DC.OperatingSystem
                    'Site'               = $DC.Site
                    'Ipv4 Address'       = $DC.Ipv4Address
                    'Ipv6 Address'       = $DC.Ipv6Address
                    'Is Global Catalog?' = $DC.IsGlobalCatalog
                    'Is Read Only?'      = $DC.IsReadOnly
                    'Ldap Port'          = $DC.LdapPort
                    'SSL Port'           = $DC.SSLPort
                }
            }
            return Format-TransposeTable $DCs
        }
    }
    return $Data
}
function Get-DomainSummary {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot,
        $Domain
    )

    $ForestName = $($ActiveDirectorySnapshot.ForestInformation.Name)
    $ForestNameDN = $($ActiveDirectorySnapshot.RootDSE.defaultNamingContext)
    $DomainNetBios = $ActiveDirectorySnapshot.DomainInformation.NetBIOSName
    $DomainName = $ActiveDirectorySnapshot.DomainInformation.DNSRoot
    $DomainDistinguishedName = $ActiveDirectorySnapshot.DomainInformation.DistinguishedName

    $Text = "Following domains exists within forest $ForestName"
    $Text0 = "Domain $DomainDistinguishedName"
    $Text1 = "Name for fully qualified domain name (FQDN): $DomainName"
    $Text2 = "Name for NetBIOS: $DomainNetBios"

    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text

    $ListDomainInformation = $null
    $ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel 0 -ListItemType Bulleted -ListValue $Text0
    $ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel 1 -ListItemType Bulleted -ListValue $Text1
    $ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel 1 -ListItemType Bulleted -ListValue $Text2
    Add-WordListItem -WordDocument $WordDocument -Paragraph $Paragraph -List $ListDomainInformation -Supress $true
}