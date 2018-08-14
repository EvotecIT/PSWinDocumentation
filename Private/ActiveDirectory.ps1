function Get-WinADForestInformation {
    [CmdletBinding()]
    param (
        [Object] $TypesRequired,
        [switch] $RequireTypes
    )
    if ($TypesRequired -eq $null) {
        Write-Verbose ' Get-WinADForestInformation - TypesRequired is null. Getting all.'
        $TypesRequired = Get-Types
    } # Gets all types

    $Data = [ordered] @{}
    $Data.Forest = $(Get-ADForest)
    $Data.RootDSE = $(Get-ADRootDSE -Properties *)
    $Data.ForestName = $Data.Forest.Name
    $Data.ForestNameDN = $Data.RootDSE.defaultNamingContext
    $Data.Domains = $Data.Forest.Domains

    if ($TypesRequired -contains [ActiveDirectory]::ForestSites -or $TypesRequired -contains [ActiveDirectory]::ForestSites1 -or $TypesRequired -contains [ActiveDirectory]::ForestSites2) {
        $Data.ForestSites = $(Get-ADReplicationSite -Filter * -Properties * )

        $Data.ForestSites1 = Invoke-Command -ScriptBlock {
            $ReturnData = @()
            foreach ($Sites in $Data.ForestSites) {
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
    if ($TypesRequired -contains [ActiveDirectory]::ForestSubnets -or $TypesRequired -contains [ActiveDirectory]::ForestSubnets1 -or $TypesRequired -contains [ActiveDirectory]::ForestSubnets2) {
        $Data.ForestSubnets = $(Get-ADReplicationSubnet -Filter * -Properties * | `
                Select-Object  Name, DisplayName, Description, Site, ProtectedFromAccidentalDeletion, Created, Modified, Deleted )
        $Data.ForestSubnets1 = Invoke-Command -ScriptBlock {
            $ReturnData = @()
            foreach ($Subnets in $Data.ForestSubnets) {
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
    if ($TypesRequired -contains [ActiveDirectory]::ForestSiteLinks) {
        $Data.ForestSiteLinks = $(
            Get-ADReplicationSiteLink -Filter * -Properties `
                Name, Cost, ReplicationFrequencyInMinutes, replInterval, ReplicationSchedule, Created, Modified, Deleted, IsDeleted, ProtectedFromAccidentalDeletion | `
                Select-Object Name, Cost, ReplicationFrequencyInMinutes, ReplInterval, Modified
        )
    }
    if ($TypesRequired -contains [ActiveDirectory]::ForestForestInformation) {
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
    if ($TypesRequired -contains [ActiveDirectory]::ForestUPNSuffixes) {
        $Data.ForestUPNSuffixes = Invoke-Command -ScriptBlock {
            $UPNSuffixList = @()
            $UPNSuffixList += $Data.Forest.RootDomain + ' (Primary / Default UPN)'
            $UPNSuffixList += $Data.Forest.UPNSuffixes
            return $UPNSuffixList
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::ForestGlobalCatalogs) {
        $Data.ForestGlobalCatalogs = $Data.Forest.GlobalCatalogs
    }
    if ($TypesRequired -contains [ActiveDirectory]::ForestSPNSuffixes) {
        $Data.ForestSPNSuffixes = $Data.Forest.SPNSuffixes
    }
    if ($TypesRequired -contains [ActiveDirectory]::ForestFSMO) {
        $Data.ForestFSMO = Invoke-Command -ScriptBlock {
            $FSMO = [ordered] @{
                'Domain Naming Master' = $Data.Forest.DomainNamingMaster
                'Schema Master'        = $Data.Forest.SchemaMaster
            }
            return $FSMO
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::ForestOptionalFeatures) {
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
    $Data.DomainRootDSE = $(Get-ADRootDSE -Server $Domain)
    $Data.DomainInformation = $(Get-ADDomain -Server $Domain)

    if ($TypesRequired -contains [ActiveDirectory]::DomainGUIDS) {
        $Data.DomainGUIDS = Invoke-Command -ScriptBlock {
            $GUID = @{}
            Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID | ForEach-Object {
                if ($GUID.Keys -notcontains $_.schemaIDGUID ) {
                    $GUID.add([System.GUID]$_.schemaIDGUID, $_.name)
                }
            }
            Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID | ForEach-Object {
                if ($GUID.Keys -notcontains $_.rightsGUID ) {
                    $GUID.add([System.GUID]$_.rightsGUID, $_.name)
                }
            }
            return $GUID
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainAuthenticationPolicies) {
        $Data.DomainAuthenticationPolicies = $(Get-ADAuthenticationPolicy -Server $Domain -LDAPFilter '(name=AuthenticationPolicy*)')
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainAuthenticationPolicySilos) {
        $Data.DomainAuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Server $Domain -Filter 'Name -like "*AuthenticationPolicySilo*"')
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainCentralAccessPolicies) {
        $Data.DomainCentralAccessPolicies = $(Get-ADCentralAccessPolicy -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainCentralAccessRules) {
        $Data.DomainCentralAccessRules = $(Get-ADCentralAccessRule -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainClaimTransformPolicies) {
        $Data.DomainClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainClaimTypes) {
        $Data.DomainClaimTypes = $(Get-ADClaimType -Server $Domain -Filter * )
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainDNSSRV -or $TypesRequired -contains [ActiveDirectory]::DomainDNSA) {
        $Data.DomainDNSData = Invoke-Command -ScriptBlock {
            $DnsSrv = @()
            $DnsA = @()

            $DnsRecords = "_kerberos._tcp.$Domain", "_ldap._tcp.$Domain"
            foreach ($DnsRecord in $DnsRecords) {
                $Value = Resolve-DnsName -Name $DnsRecord -Type SRV | Select *
                foreach ($V in $Value) {
                    if ($V.QueryType -eq 'SRV') {
                        $DnsSrv += $V
                    } else {
                        $DnsA += $V
                    }
                }
            }
            $ReturnData = @{
                # QueryType, Target, NameTarget, Priority, Weight, Port, Name, Type, CharacterSet, Section
                SRV = $DnsSrv | Select-Object Target, NameTarget, Priority, Weight, Port, Name # Type, QueryType, CharacterSet, Section
                # Address, IPAddress, QueryType, IP4Address, Name, Type, CharacterSet, Section, DataLength, TTL
                A   = $DnsA | Select-Object Address, IPAddress, IP4Address, Name, Type, DataLength, TTL # QueryType, CharacterSet, Section
            }
            return $ReturnData
        }
        $Data.DomainDNSSrv = $Data.DomainDNSData.SRV
        $Data.DomainDNSA = $Data.DomainDNSData.A
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainFSMO -or $TypesRequired -contains [ActiveDirectory]::DomainDomainTrusts) {
        # required for multiple use cases FSMO/DomainTrusts
        $Data.DomainFSMO = [ordered] @{
            'PDC Emulator'          = $Data.DomainInformation.PDCEmulator
            'RID Master'            = $Data.DomainInformation.RIDMaster
            'Infrastructure Master' = $Data.DomainInformation.InfrastructureMaster
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainTrusts) {
        ## requires both DomainTrusts and FSMO.
        $Data.DomainTrustsClean = (Get-ADTrust -Server $Domain -Filter * -Properties *)
        $Data.DomainTrusts = Invoke-Command -ScriptBlock {
            $DomainPDC = $Data.DomainFSMO.'PDC Emulator'
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
    if ($TypesRequired -contains [ActiveDirectory]::DomainGroupPolicies -or $TypesRequired -contains [ActiveDirectory]::DomainGroupPoliciesDetails -or $TypesRequired -contains [ActiveDirectory]::DomainGroupPoliciesACL) {
        $Data.DomainGroupPoliciesClean = $(Get-GPO -Domain $Domain -All)
        $Data.DomainGroupPolicies = Invoke-Command -ScriptBlock {
            $GroupPolicies = @()
            foreach ($gpo in $Data.DomainGroupPoliciesClean) {
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
        $Data.DomainGroupPoliciesDetails = Invoke-Command -ScriptBlock {
            Write-Verbose -Message "Get-WinADDomainInformation - Group Policies Details"
            $Output = @()
            ForEach ($GPO in $Data.DomainGroupPoliciesClean) {
                [xml]$XmlGPReport = $GPO.generatereport('xml')
                #GPO version
                if ($XmlGPReport.GPO.Computer.VersionDirectory -eq 0 -and $XmlGPReport.GPO.Computer.VersionSysvol -eq 0) {$ComputerSettings = "NeverModified"}else {$ComputerSettings = "Modified"}
                if ($XmlGPReport.GPO.User.VersionDirectory -eq 0 -and $XmlGPReport.GPO.User.VersionSysvol -eq 0) {$UserSettings = "NeverModified"}else {$UserSettings = "Modified"}
                #GPO content
                if ($XmlGPReport.GPO.User.ExtensionData -eq $null) {$UserSettingsConfigured = $false}else {$UserSettingsConfigured = $true}
                if ($XmlGPReport.GPO.Computer.ExtensionData -eq $null) {$ComputerSettingsConfigured = $false}else {$ComputerSettingsConfigured = $true}
                #Output
                $Output += [ordered] @{
                    'Name'                   = $XmlGPReport.GPO.Name
                    'Links'                  = $XmlGPReport.GPO.LinksTo | Select-Object -ExpandProperty SOMPath
                    'Has Computer Settings'  = $ComputerSettingsConfigured
                    'Has User Settings'      = $UserSettingsConfigured
                    'User Enabled'           = $XmlGPReport.GPO.User.Enabled
                    'Computer Enabled'       = $XmlGPReport.GPO.Computer.Enabled
                    'Computer Settings'      = $ComputerSettings
                    'User Settings'          = $UserSettings
                    'Gpo Status'             = $GPO.GpoStatus
                    'Creation Time'          = $GPO.CreationTime
                    'Modification Time'      = $GPO.ModificationTime
                    'WMI Filter'             = $GPO.WmiFilter.name
                    'WMI Filter Description' = $GPO.WmiFilter.Description
                    'Path'                   = $GPO.Path
                    'GUID'                   = $GPO.Id
                    'SDDL'                   = $XmlGPReport.GPO.SecurityDescriptor.SDDL.'#text'
                    #'ACLs'                   = $XmlGPReport.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object -Process {
                    #    New-Object -TypeName PSObject -Property @{
                    #        'User'            = $_.trustee.name.'#Text'
                    #        'Permission Type' = $_.type.PermissionType
                    #        'Inherited'       = $_.Inherited
                    #        'Permissions'     = $_.Standard.GPOGroupedAccessEnum
                    #    }
                    #}
                }
            }
            return Format-TransposeTable $Output
        }
        $Data.DomainGroupPoliciesACL = Invoke-Command -ScriptBlock {
            Write-Verbose -Message "Get-WinADDomainInformation - Group Policies ACLs"
            $Output = @()
            ForEach ($GPO in $Data.DomainGroupPoliciesClean) {
                [xml]$XmlGPReport = $GPO.generatereport('xml')
                $ACLs = $XmlGPReport.GPO.SecurityDescriptor.Permissions.TrusteePermissions
                foreach ($ACL in $ACLS) {
                    $Output += [ordered] @{
                        'GPO Name'        = $GPO.DisplayName
                        'User'            = $ACL.trustee.name.'#Text'
                        'Permission Type' = $ACL.type.PermissionType
                        'Inherited'       = $ACL.Inherited
                        'Permissions'     = $ACL.Standard.GPOGroupedAccessEnum
                    }
                }
            }
            return Format-TransposeTable $Output
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainDefaultPasswordPolicy) {
        $Data.DomainDefaultPasswordPolicy = Invoke-Command -ScriptBlock {
            $Policy = $(Get-ADDefaultDomainPasswordPolicy -Server $Domain)
            $Data = [ordered] @{
                'Complexity Enabled'            = $Policy.ComplexityEnabled
                'Lockout Duration'              = $Policy.LockoutDuration
                'Lockout Observation Window'    = $Policy.LockoutObservationWindow
                'Lockout Threshold'             = $Policy.LockoutThreshold
                'Max Password Age'              = $Policy.MaxPasswordAge
                'Min Password Length'           = $Policy.MinPasswordLength
                'Min Password Age'              = $Policy.MinPasswordAge
                'Password History Count'        = $Policy.PasswordHistoryCount
                'Reversible Encryption Enabled' = $Policy.ReversibleEncryptionEnabled
                'Distinguished Name'            = $Policy.DistinguishedName
            }
            return $Data
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainPriviligedGroupMembers) {
        Write-Verbose "Get-WinADDomainInformation - TypesRequired: PriviligedGroupMembers"
        $Data.DomainPriviligedGroupMembers = Get-PrivilegedGroupsMembers -Domain $Data.DomainInformation.DNSRoot -DomainSID $Data.DomainInformation.DomainSid
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainOrganizationalUnits -or $TypesRequired -contains [ActiveDirectory]::DomainContainers) {
        #CanonicalName, ManagedBy, ProtectedFromAccidentalDeletion, Created, Modified, Deleted, PostalCode, City, Country, State, StreetAddress, ProtectedFromAccidentalDeletion, DistinguishedName, ObjectGUID
        # $Data.DomainContainers = Get-ADObject -SearchBase $Data.DomainInformation.DistinguishedName -SearchScope OneLevel -LDAPFilter '(objectClass=container)' -Properties *
        $Data.DomainOrganizationalUnitsClean = $(Get-ADOrganizationalUnit -Server $Domain -Properties * -Filter * )
        $Data.DomainOrganizationalUnits = Invoke-Command -ScriptBlock {
            return $Data.DomainOrganizationalUnitsClean | Select-Object `
            @{ n = 'Canonical Name'; e = { $_.CanonicalName }},
            @{ n = 'Managed By'; e = { $_.ManagedBy }},
            @{ n = 'Protected'; e = { $_.ProtectedFromAccidentalDeletion }},
            Created,
            Modified,
            Deleted,
            @{ n = 'Postal Code'; e = { $_.PostalCode }},
            City,
            Country,
            State,
            @{ n = 'Street Address'; e = { $_.StreetAddress }},
            DistinguishedName,
            ObjectGUID | Sort-Object CanonicalName
        }
        $Data.DomainOrganizationalUnitsDN = Invoke-Command -ScriptBlock {
            $OUs = @()
            $OUs += $Data.DomainInformation.DistinguishedName
            $OUS += $Data.DomainOrganizationalUnitsClean.DistinguishedName
            $OUs += $Data.DomainContainers.DistinguishedName
            return $OUs
        }
        $Data.DomainOrganizationalUnitsACL = Invoke-Command -ScriptBlock {
            $ReportBasic = @()
            $ReportExtented = @()
            $OUs = @()
            #$OUs += @{ Name = 'Root'; Value = $Data.DomainRootDSE.rootDomainNamingContext }
            foreach ($OU in $Data.DomainOrganizationalUnitsClean) {
                $OUs += @{ Name = 'Organizational Unit'; Value = $OU.DistinguishedName }
                Write-Verbose "1. $($Ou.DistinguishedName)"
            }
            #foreach ($OU in $Data.DomainContainers) {
            #    $OUs += @{ Name = 'Container'; Value = $OU.DistinguishedName }
            #    Write-Verbose "2. $($Ou.DistinguishedName)"
            #}
            $PSDriveName = $Data.DomainInformation.NetBIOSName
            New-PSDrive -Name $PSDriveName -Root "" -PsProvider ActiveDirectory -Server $Domain

            ForEach ($OU in $OUs) {
                Write-Verbose "3. $($Ou.Value)"
                $ReportBasic += Get-Acl -Path "$PSDriveName`:\$($OU.Value)" | Select-Object `
                @{name = 'Distinguished Name'; expression = { $OU.Value}},
                @{name = 'Type'; expression = { $OU.Name }},
                @{name = 'Owner'; expression = {$_.Owner}},
                @{name = 'Group'; expression = {$_.Group}},
                @{name = 'Are AccessRules Protected'; expression = { $_.AreAccessRulesProtected}},
                @{name = 'Are AuditRules Protected'; expression = {$_.AreAuditRulesProtected}},
                @{name = 'Are AccessRules Canonical'; expression = { $_.AreAccessRulesCanonical}},
                @{name = 'Are AuditRules Canonical'; expression = { $_.AreAuditRulesCanonical}},
                @{name = 'Sddl'; expression = {$_.Sddl}}

                $ReportExtented += Get-Acl -Path "$PSDriveName`:\$($OU.Value)" | `
                    Select-Object -ExpandProperty Access | `
                    Select-Object `
                @{name = 'Distinguished Name'; expression = {$OU.Value}},
                @{name = 'Type'; expression = {$OU.Name}},
                @{name = 'AccessControlType'; expression = {$_.AccessControlType }},
                @{name = 'ObjectType Name'; expression = {if ($_.objectType.ToString() -eq '00000000-0000-0000-0000-000000000000') {'All'} Else {$GUID.Item($_.objectType)}}},
                @{name = 'Inherited ObjectType Name'; expression = {$GUID.Item($_.inheritedObjectType)}},
                @{name = 'ActiveDirectoryRights'; expression = {$_.ActiveDirectoryRights}},
                @{name = 'InheritanceType'; expression = {$_.InheritanceType}},
                @{name = 'ObjectType'; expression = {$_.ObjectType}},
                @{name = 'InheritedObjectType'; expression = {$_.InheritedObjectType}},
                @{name = 'ObjectFlags'; expression = {$_.ObjectFlags}},
                @{name = 'IdentityReference'; expression = {$_.IdentityReference}},
                @{name = 'IsInherited'; expression = {$_.IsInherited}},
                @{name = 'InheritanceFlags'; expression = {$_.InheritanceFlags}},
                @{name = 'PropagationFlags'; expression = {$_.PropagationFlags}}


            }
            return @{ Basic = $ReportBasic; Extended = $ReportExtented }
        }
        $Data.DomainOrganizationalUnitsBasicACL = $Data.DomainOrganizationalUnitsACL.Basic
        $Data.DomainOrganizationalUnitsExtended = $Data.DomainOrganizationalUnitsACL.Extended
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainAdministrators) {
        $Data.DomainAdministratorsClean = $( Get-ADGroup -Server $Domain -Identity $('{0}-512' -f $Data.DomainInformation.DomainSID) | Get-ADGroupMember -Server $Domain -Recursive | Get-ADUser -Server $Domain)
        $Data.DomainAdministrators = $Data.DomainAdministratorsClean | Select-Object Name, SamAccountName, UserPrincipalName, Enabled
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainUsers -or $TypesRequired -contains [ActiveDirectory]::DomainUsersCount) {
        Write-Verbose 'Get-WinDomainInformation - Getting All Users'
        $Data.DomainUsers = Invoke-Command -ScriptBlock {
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
    if ($TypesRequired -contains [ActiveDirectory]::DomainUsersCount) {
        Write-Verbose 'Get-WinDomainInformation - Getting All Users Count'
        $Data.DomainUsersCount = [ordered] @{
            'Users Count Incl. System'            = Get-ObjectCount -Object $Data.DomainUsers.Users
            'Users Count'                         = Get-ObjectCount -Object $Data.DomainUsers.UsersAll
            'Users Expired'                       = Get-ObjectCount -Object $Data.DomainUsers.UsersExpiredExclDisabled
            'Users Expired Incl. Disabled'        = Get-ObjectCount -Object $Data.DomainUsers.UsersExpiredInclDisabled
            'Users Never Expiring'                = Get-ObjectCount -Object $Data.DomainUsers.UsersNeverExpiring
            'Users Never Expiring Incl. Disabled' = Get-ObjectCount -Object $Data.DomainUsers.UsersNeverExpiringInclDisabled
            'Users System Accounts'               = Get-ObjectCount -Object $Data.DomainUsers.UsersSystemAccounts
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainDomainControllers) {
        $Data.DomainControllersClean = $(Get-ADDomainController -Server $Domain -Filter * )
        $Data.DomainControllers = Invoke-Command -ScriptBlock {
            $DCs = @()
            foreach ($Policy in $Data.DomainControllersClean) {
                $DCs += [ordered] @{
                    'Name'               = $Policy.Name
                    'Host Name'          = $Policy.HostName
                    'Operating System'   = $Policy.OperatingSystem
                    'Site'               = $Policy.Site
                    'Ipv4 Address'       = $Policy.Ipv4Address
                    'Ipv6 Address'       = $Policy.Ipv6Address
                    'Is Global Catalog?' = $Policy.IsGlobalCatalog
                    'Is Read Only?'      = $Policy.IsReadOnly
                    'Ldap Port'          = $Policy.LdapPort
                    'SSL Port'           = $Policy.SSLPort
                }
            }
            return Format-TransposeTable $DCs
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainFineGrainedPolicies) {
        <#

        AppliesTo                   : {CN=GDS-FineGrainedPolicy-Test,OU=Groups,OU=Production,DC=ad,DC=evotec,DC=pl}
        ComplexityEnabled           : False
        DistinguishedName           : CN=Fine Policy Test,CN=Password Settings Container,CN=System,DC=ad,DC=evotec,DC=pl
        LockoutDuration             : 00:30:00
        LockoutObservationWindow    : 00:30:00
        LockoutThreshold            : 0
        MaxPasswordAge              : 00:00:00
        MinPasswordAge              : 00:00:00
        MinPasswordLength           : 4
        Name                        : Fine Policy Test
        ObjectClass                 : msDS-PasswordSettings
        ObjectGUID                  : db28647d-d5c1-45b0-8671-4b56228e0c18
        PasswordHistoryCount        : 0
        Precedence                  : 200
        ReversibleEncryptionEnabled : True
        #>
        $Data.FineGrainedPolicies = Invoke-Command -ScriptBlock {
            $FineGrainedPoliciesData = Get-ADFineGrainedPasswordPolicy -Filter * -Server $Domain
            $FineGrainedPolicies = @()
            foreach ($Policy in $FineGrainedPoliciesData) {
                $FineGrainedPolicies += [ordered] @{
                    'Name'                          = $Policy.Name
                    'Complexity Enabled'            = $Policy.ComplexityEnabled
                    'Lockout Duration'              = $Policy.LockoutDuration
                    'Lockout Observation Window'    = $Policy.LockoutObservationWindow
                    'Lockout Threshold'             = $Policy.LockoutThreshold
                    'Max Password Age'              = $Policy.MaxPasswordAge
                    'Min Password Length'           = $Policy.MinPasswordLength
                    'Min Password Age'              = $Policy.MinPasswordAge
                    'Password History Count'        = $Policy.PasswordHistoryCount
                    'Reversible Encryption Enabled' = $Policy.ReversibleEncryptionEnabled
                    'Precedence'                    = $Policy.Precedence
                    'Applies To'                    = $Policy.AppliesTo # get all groups / usrs and convert to data TODO
                    'Distinguished Name'            = $Policy.DistinguishedName
                }
            }
            return Format-TransposeTable $FineGrainedPolicies
        }
    }
    return $Data
}