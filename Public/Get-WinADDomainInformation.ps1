function Get-WinADDomainInformation {
    [CmdletBinding()]
    param (
        [string] $Domain,
        [Object] $TypesRequired,
        [string] $PathToPasswords,
        [string] $PathToPasswordsHashes
    )
    if ([string]::IsNullOrEmpty($Domain)) {
        Write-Warning 'Get-WinADDomainInformation - $Domain parameter is empty. Try your domain name like ad.evotec.xyz. Skipping for now...'
        return
    }
    if ($TypesRequired -eq $null) {
        Write-Verbose 'Get-WinADDomainInformation - TypesRequired is null. Getting all.'
        $TypesRequired = Get-Types -Types ([ActiveDirectory])
    } # Gets all types
    $Data = [ordered] @{}
    Write-Verbose "Getting domain information - $Domain DomainRootDSE"
    $Data.DomainRootDSE = $(Get-ADRootDSE -Server $Domain)
    Write-Verbose "Getting domain information - $Domain DomainInformation"
    $Data.DomainInformation = $(Get-ADDomain -Server $Domain)
    Write-Verbose "Getting domain information - $Domain DomainGroupsFullList"
    $Data.DomainGroupsFullList = Get-ADGroup -Server $Domain -Filter * -ResultPageSize 500000 -Properties * | Select-Object * -ExcludeProperty *Certificate, PropertyNames, *Properties, PropertyCount, Certificates, nTSecurityDescriptor
    Write-Verbose "Getting domain information - $Domain DomainUsersFullList"
    $Data.DomainUsersFullList = Get-ADUser -Server $Domain -ResultPageSize 500000 -Filter * -Properties *, "msDS-UserPasswordExpiryTimeComputed" | Select-Object * -ExcludeProperty *Certificate, PropertyNames, *Properties, PropertyCount, Certificates, nTSecurityDescriptor
    Write-Verbose "Getting domain information - $Domain DomainComputersFullList"
    $Data.DomainComputersFullList = Get-ADComputer -Server $Domain -Filter * -ResultPageSize 500000 -Properties * | Select-Object * -ExcludeProperty *Certificate, PropertyNames, *Properties, PropertyCount, Certificates, nTSecurityDescriptor

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [ActiveDirectory]::DomainComputersAll,
            [ActiveDirectory]::DomainComputersAllCount,
            [ActiveDirectory]::DomainServers,
            [ActiveDirectory]::DomainServersCount,
            [ActiveDirectory]::DomainComputers,
            [ActiveDirectory]::DomainComputersCount,
            [ActiveDirectory]::DomainComputersUnknown,
            [ActiveDirectory]::DomainComputersUnknownCount
        )) {
        $Data.DomainComputersAll = $Data.DomainComputersFullList  | Select-Object Name, SamAccountName, Enabled, PasswordLastSet, IPv4Address, IPv6Address, DNSHostName, ManagedBy, OperatingSystem*, PasswordNeverExpires, PasswordNotRequired, UserPrincipalName, LastLogonDate, LockedOut, LogonCount, CanonicalName, SID, Created, Modified, Deleted, MemberOf
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainComputersAllCount) {
        $Data.DomainComputersAllCount = $Data.DomainComputersAll | Group-Object -Property OperatingSystem | Select-Object @{ L = 'System Name'; Expression = { if ($_.Name -ne '') { $_.Name } else { 'Unknown' } }} , @{ L = 'System Count'; Expression = { $_.Count }}
    }

    if ($TypesRequired -contains [ActiveDirectory]::DomainServers) {
        $Data.DomainServers = $Data.DomainComputersAll | Where-Object { $_.OperatingSystem -like 'Windows Server*' }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainServersCount) {
        $Data.DomainServersCount = $Data.DomainServers | Group-Object -Property OperatingSystem | Select-Object @{ L = 'System Name'; Expression = { if ($_.Name -ne '') { $_.Name } else { 'N/A' } }} , @{ L = 'System Count'; Expression = { $_.Count }}
    }

    if ($TypesRequired -contains [ActiveDirectory]::DomainComputers) {
        $Data.DomainComputers = $Data.DomainComputersAll | Where-Object { $_.OperatingSystem -notlike 'Windows Server*' -and $_.OperatingSystem -ne $null }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainComputersCount) {
        $Data.DomainComputersCount = $Data.DomainComputers | Group-Object -Property OperatingSystem | Select-Object @{ L = 'System Name'; Expression = { if ($_.Name -ne '') { $_.Name } else { 'N/A' } }} , @{ L = 'System Count'; Expression = { $_.Count }}
    }

    if ($TypesRequired -contains [ActiveDirectory]::DomainComputersUnknown) {
        $Data.DomainComputersUnknown = $Data.DomainComputersAll | Where-Object { $_.OperatingSystem -eq $null }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainComputersUnknownCount) {
        $Data.DomainComputersUnknownCount = $Data.DomainComputersUnknown | Group-Object -Property OperatingSystem | Select-Object @{ L = 'System Name'; Expression = { if ($_.Name -ne '') { $_.Name } else { 'Unknown' } }} , @{ L = 'System Count'; Expression = { $_.Count }}
    }

    if ($TypesRequired -contains [ActiveDirectory]::DomainRIDs) {
        # Critical for RID Pool Depletion: https://blogs.technet.microsoft.com/askds/2011/09/12/managing-rid-pool-depletion/
        $Data.DomainRIDs = Invoke-Command -ScriptBlock {
            Write-Verbose "Getting domain information - $Domain DomainRIDs"
            #Write-Verbose "Get-WinADDomainInformation - RID Master: $($Data.DomainInformation.RIDMaster) - DN: $($Data.DomainInformation.DistinguishedName)"
            $rID = [ordered] @{}
            $rID.'rIDs Master' = $Data.DomainInformation.RIDMaster

            $property = get-adobject "cn=rid manager$,cn=system,$($Data.DomainInformation.DistinguishedName)" -property RidAvailablePool -Server $rID.'rIDs Master'
            [int32]$totalSIDS = $($property.RidAvailablePool) / ([math]::Pow(2, 32))
            [int64]$temp64val = $totalSIDS * ([math]::Pow(2, 32))
            [int32]$currentRIDPoolCount = $($property.RidAvailablePool) - $temp64val
            [int64]$RidsRemaining = $totalSIDS - $currentRIDPoolCount

            $Rid.'rIDs Available Pool' = $property.RidAvailablePool
            $rID.'rIDs Total SIDs' = $totalSIDS
            $rID.'rIDs Issued' = $CurrentRIDPoolCount
            $rID.'rIDs Remaining' = $RidsRemaining
            $rID.'rIDs Percentage' = if ($RidsRemaining -eq 0) { $RidsRemaining.ToString("P") } else { ($currentRIDPoolCount / $RidsRemaining * 100).ToString("P") }
            return $rID
        }
    }
    if ($TypesRequired -contains [ActiveDirectory]::DomainGUIDS) {
        Write-Verbose "Getting domain information - $Domain DomainGUIDS"
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
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainAuthenticationPolicies)) {
        Write-Verbose "Getting domain information - $Domain DomainAuthenticationPolicies"
        $Data.DomainAuthenticationPolicies = $(Get-ADAuthenticationPolicy -Server $Domain -LDAPFilter '(name=AuthenticationPolicy*)')
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainAuthenticationPolicySilos)) {
        Write-Verbose "Getting domain information - $Domain DomainAuthenticationPolicySilos"
        $Data.DomainAuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Server $Domain -Filter 'Name -like "*AuthenticationPolicySilo*"')
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainCentralAccessPolicies)) {
        Write-Verbose "Getting domain information - $Domain DomainCentralAccessPolicies"
        $Data.DomainCentralAccessPolicies = $(Get-ADCentralAccessPolicy -Server $Domain -Filter * )
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainCentralAccessRules)) {
        Write-Verbose "Getting domain information - $Domain DomainCentralAccessRules"
        $Data.DomainCentralAccessRules = $(Get-ADCentralAccessRule -Server $Domain -Filter * )
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainClaimTransformPolicies)) {
        Write-Verbose "Getting domain information - $Domain DomainClaimTransformPolicies"
        $Data.DomainClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Server $Domain -Filter * )
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainClaimTypes)) {
        Write-Verbose "Getting domain information - $Domain DomainClaimTypes"
        $Data.DomainClaimTypes = $(Get-ADClaimType -Server $Domain -Filter * )
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainDNSSRV, [ActiveDirectory]::DomainDNSA )) {
        Write-Verbose "Getting domain information - $Domain DomainDNSSRV / DomainDNSA"
        $Data.DomainDNSData = Invoke-Command -ScriptBlock {
            $DnsSrv = @()
            $DnsA = @()

            $DnsRecords = "_kerberos._tcp.$Domain", "_ldap._tcp.$Domain"
            foreach ($DnsRecord in $DnsRecords) {
                $Value = Resolve-DnsName -Name $DnsRecord -Type SRV -Verbose:$false -ErrorAction SilentlyContinue | Select *
                if ($Value -eq $null) { Write-Warning 'Getting domain information - DomainDNSSRV / DomainDNSA - Failed!'}
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
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainFSMO, [ActiveDirectory]::DomainTrusts )) {
        Write-Verbose "Getting domain information - $Domain DomainFSMO"
        # required for multiple use cases FSMO/DomainTrusts
        $Data.DomainFSMO = [ordered] @{
            'PDC Emulator'          = $Data.DomainInformation.PDCEmulator
            'RID Master'            = $Data.DomainInformation.RIDMaster
            'Infrastructure Master' = $Data.DomainInformation.InfrastructureMaster
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainTrusts)) {
        Write-Verbose "Getting domain information - $Domain DomainTrusts"
        ## requires both DomainTrusts and FSMO.
        $Data.DomainTrustsClean = (Get-ADTrust -Server $Domain -Filter * -Properties *)
        $Data.DomainTrusts = Invoke-Command -ScriptBlock {
            $DomainPDC = $Data.DomainFSMO.'PDC Emulator'
            $Trust = $Data.DomainTrustsClean

            $TrustWMI = Get-CimInstance -ClassName Microsoft_DomainTrustStatus -Namespace root\MicrosoftActiveDirectory -ComputerName $DomainPDC -ErrorAction SilentlyContinue -Verbose:$false | Select-Object TrustIsOK, TrustStatus, TrustStatusString, PSComputerName, TrustedDCName

            if ($Trust) {
                $ReturnData = [ordered] @{
                    'Trust Source'               = $Domain
                    'Trust Target'               = $Trust.Target
                    'Trust Direction'            = $Trust.Direction
                    'Trust Attributes'           = Set-TrustAttributes -Value $Trust.TrustAttributes
                    #'Trust OK'                   = $TrustWMI.TrustIsOK
                    #'Trust Status'               = $TrustWMI.TrustStatus
                    'Trust Status'               = if ($TrustWMI -ne $null) { $TrustWMI.TrustStatusString } else { 'N/A' }
                    'Forest Transitive'          = $Trust.ForestTransitive
                    'Selective Authentication'   = $Trust.SelectiveAuthentication
                    'SID Filtering Forest Aware' = $Trust.SIDFilteringForestAware
                    'SID Filtering Quarantined'  = $Trust.SIDFilteringQuarantined
                    'Disallow Transivity'        = $Trust.DisallowTransivity
                    'Intra Forest'               = $Trust.IntraForest
                    'Tree Parent?'               = $Trust.IsTreeParent
                    'Tree Root?'                 = $Trust.IsTreeRoot
                    'TGTDelegation'              = $Trust.TGTDelegation
                    'TrustedPolicy'              = $Trust.TrustedPolicy
                    'TrustingPolicy'             = $Trust.TrustingPolicy
                    'TrustType'                  = $Trust.TrustType
                    'UplevelOnly'                = $Trust.UplevelOnly
                    'UsesAESKeys'                = $Trust.UsesAESKeys
                    'UsesRC4Encryption'          = $Trust.UsesRC4Encryption
                    'Trust Source DC'            = if ($TrustWMI -ne $null) { $TrustWMI.PSComputerName } else { 'N/A' }
                    'Trust Target DC'            = if ($TrustWMI -ne $null) { $TrustWMI.TrustedDCName.Replace('\\', '') } else { 'N/A'}
                    'Trust Source DN'            = $Trust.Source
                    'ObjectGUID'                 = $Trust.ObjectGUID
                    'Created'                    = $Trust.Created
                    'Modified'                   = $Trust.Modified
                    'Deleted'                    = $Trust.Deleted
                    'SID'                        = $Trust.securityIdentifier
                }
            } else {
                $ReturnData = $null
            }
            return Format-TransposeTable $ReturnData
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [ActiveDirectory]::DomainGroupPolicies,
            [ActiveDirectory]::DomainGroupPoliciesDetails,
            [ActiveDirectory]::DomainGroupPoliciesACL
        )) {
        Write-Verbose "Getting domain information - $Domain DomainGroupPolicies"
        $Data.DomainGroupPoliciesClean = $(Get-GPO -Domain $Domain -Server $Domain -All)
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
            Write-Verbose -Message "Getting domain information - $Domain Group Policies Details"
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
            Write-Verbose -Message "Getting domain information - $Domain Group Policies ACLs"
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
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainDefaultPasswordPolicy)) {
        Write-Verbose -Message "Getting domain information - $Domain DomainDefaultPasswordPolicy"
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
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [ActiveDirectory]::DomainOrganizationalUnits,
            [ActiveDirectory]::DomainContainers,
            [ActiveDirectory]::DomainOrganizationalUnitsDN,
            [ActiveDirectory]::DomainOrganizationalUnitsACL,
            [ActiveDirectory]::DomainOrganizationalUnitsBasicACL,
            [ActiveDirectory]::DomainOrganizationalUnitsExtended
        )) {
        Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnits"
        $Data.DomainOrganizationalUnitsClean = $(Get-ADOrganizationalUnit -Server $Domain -Properties * -Filter * )
        $Data.DomainOrganizationalUnits = Invoke-Command -ScriptBlock {
            return $Data.DomainOrganizationalUnitsClean | Select-Object `
            @{ n = 'Canonical Name'; e = { $_.CanonicalName }},
            @{ n = 'Managed By'; e = {
                    (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $_.ManagedBy -Verbose).Name
                }
            },
            @{ n = 'Manager Email'; e = {
                    (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $_.ManagedBy -Verbose).EmailAddress
                }
            },
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
            ObjectGUID | Sort-Object 'Canonical Name'
        }
        Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsDN"
        $Data.DomainOrganizationalUnitsDN = Invoke-Command -ScriptBlock {
            $OUs = @()
            $OUs += $Data.DomainInformation.DistinguishedName
            $OUS += $Data.DomainOrganizationalUnitsClean.DistinguishedName
            $OUs += $Data.DomainContainers.DistinguishedName
            return $OUs
        }
        Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsACL"
        $Data.DomainOrganizationalUnitsACL = Invoke-Command -ScriptBlock {
            $ReportBasic = @()
            $ReportExtented = @()
            $OUs = @()
            #$OUs += @{ Name = 'Root'; Value = $Data.DomainRootDSE.rootDomainNamingContext }
            foreach ($OU in $Data.DomainOrganizationalUnitsClean) {
                $OUs += @{ Name = 'Organizational Unit'; Value = $OU.DistinguishedName }
                #Write-Verbose "1. $($Ou.DistinguishedName)"
            }
            #foreach ($OU in $Data.DomainContainers) {
            #    $OUs += @{ Name = 'Container'; Value = $OU.DistinguishedName }
            #    Write-Verbose "2. $($Ou.DistinguishedName)"
            #}
            $PSDriveName = $Data.DomainInformation.NetBIOSName
            New-PSDrive -Name $PSDriveName -Root "" -PsProvider ActiveDirectory -Server $Domain

            ForEach ($OU in $OUs) {
                #Write-Verbose "3. $($Ou.Value)"
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
        Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsBasicACL"
        $Data.DomainOrganizationalUnitsBasicACL = $Data.DomainOrganizationalUnitsACL.Basic
        Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsExtended"
        $Data.DomainOrganizationalUnitsExtended = $Data.DomainOrganizationalUnitsACL.Extended
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [ActiveDirectory]::DomainUsers,
            [ActiveDirectory]::DomainUsersAll,
            [ActiveDirectory]::DomainUsersSystemAccounts,
            [ActiveDirectory]::DomainUsersNeverExpiring,
            [ActiveDirectory]::DomainUsersNeverExpiringInclDisabled,
            [ActiveDirectory]::DomainUsersExpiredInclDisabled,
            [ActiveDirectory]::DomainUsersExpiredExclDisabled,
            [ActiveDirectory]::DomainUsersCount
        )) {

        $Data.DomainUsers = Invoke-Command -ScriptBlock {
            Write-Verbose "Getting domain information - $Domain DomainUsers"
            return Get-WinUsers -Users $Data.DomainUsersFullList -Domain $Domain -ADCatalog $Data.DomainUsersFullList, $Data.DomainComputersFullList, $Data.DomainGroupsFullList -ADCatalogUsers $Data.DomainUsersFullList
        }
        Write-Verbose "Getting domain information - $Domain DomainUsersAll"
        $Data.DomainUsersAll = $Data.DomainUsers | Where { $_.PasswordNotRequired -eq $False } #| Select-Object * #Name, SamAccountName, UserPrincipalName, Enabled
        Write-Verbose "Getting domain information - $Domain DomainUsersSystemAccounts"
        $Data.DomainUsersSystemAccounts = $Data.DomainUsers | Where { $_.PasswordNotRequired -eq $true } #| Select-Object * #Name, SamAccountName, UserPrincipalName, Enabled
        Write-Verbose "Getting domain information - $Domain DomainUsersNeverExpiring"
        $Data.DomainUsersNeverExpiring = $Data.DomainUsers | Where { $_.PasswordNeverExpires -eq $true -and $_.Enabled -eq $true -and $_.PasswordNotRequired -eq $false } #| Select-Object * #Name, SamAccountName, UserPrincipalName, Enabled
        Write-Verbose "Getting domain information - $Domain DomainUsersNeverExpiringInclDisabled"
        $Data.DomainUsersNeverExpiringInclDisabled = $Data.DomainUsers | Where { $_.PasswordNeverExpires -eq $true -and $_.PasswordNotRequired -eq $false } #| Select-Object * #Name, SamAccountName, UserPrincipalName, Enabled
        Write-Verbose "Getting domain information - $Domain DomainUsersExpiredInclDisabled"
        $Data.DomainUsersExpiredInclDisabled = $Data.DomainUsers | Where { $_.PasswordNeverExpires -eq $false -and $_.DaysToExpire -le 0 -and $_.PasswordNotRequired -eq $false } #| Select-Object * #Name, SamAccountName, UserPrincipalName, Enabled
        Write-Verbose "Getting domain information - $Domain DomainUsersExpiredExclDisabled"
        $Data.DomainUsersExpiredExclDisabled = $Data.DomainUsers | Where { $_.PasswordNeverExpires -eq $false -and $_.DaysToExpire -le 0 -and $_.Enabled -eq $true -and $_.PasswordNotRequired -eq $false } #| Select-Object * # Name, SamAccountName, UserPrincipalName, Enabled
        Write-Verbose "Getting domain information - $Domain All Users Count"
        $Data.DomainUsersCount = [ordered] @{
            'Users Count Incl. System'            = Get-ObjectCount -Object $Data.DomainUsers
            'Users Count'                         = Get-ObjectCount -Object $Data.DomainUsersAll
            'Users Expired'                       = Get-ObjectCount -Object $Data.DomainUsersExpiredExclDisabled
            'Users Expired Incl. Disabled'        = Get-ObjectCount -Object $Data.DomainUsersExpiredInclDisabled
            'Users Never Expiring'                = Get-ObjectCount -Object $Data.DomainUsersNeverExpiring
            'Users Never Expiring Incl. Disabled' = Get-ObjectCount -Object $Data.DomainUsersNeverExpiringInclDisabled
            'Users System Accounts'               = Get-ObjectCount -Object $Data.DomainUsersSystemAccounts
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainControllers )) {
        Write-Verbose "Getting domain information - $Domain DomainControllers"
        $Data.DomainControllersClean = $(Get-ADDomainController -Server $Domain -Filter * )
        $Data.DomainControllers = Invoke-Command -ScriptBlock {
            $DCs = @()
            foreach ($Policy in $Data.DomainControllersClean) {
                $DCs += [ordered] @{
                    'Name'             = $Policy.Name
                    'Host Name'        = $Policy.HostName
                    'Operating System' = $Policy.OperatingSystem
                    'Site'             = $Policy.Site
                    'Ipv4'             = $Policy.Ipv4Address
                    'Ipv6'             = $Policy.Ipv6Address
                    'Global Catalog?'  = $Policy.IsGlobalCatalog
                    'Read Only?'       = $Policy.IsReadOnly
                    'Ldap Port'        = $Policy.LdapPort
                    'SSL Port'         = $Policy.SSLPort
                }
            }
            return Format-TransposeTable $DCs
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainFineGrainedPolicies)) {
        Write-Verbose "Getting domain information - $Domain DomainFineGrainedPolicies"
        $Data.DomainFineGrainedPolicies = Invoke-Command -ScriptBlock {
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
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainFineGrainedPoliciesUsers)) {
        Write-Verbose "Getting domain information - $Domain DomainFineGrainedPoliciesUsers"
        $Data.DomainFineGrainedPoliciesUsers = Invoke-Command -ScriptBlock {
            $PolicyUsers = @()
            foreach ($Policy in $Data.DomainFineGrainedPolicies) {
                $Users = @()
                $Groups = @()
                foreach ($U in $Policy.'Applies To') {
                    $Users += Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $U
                    $Groups += Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainGroupsFullList -DistinguishedName $U
                }
                foreach ($User in $Users) {
                    $PolicyUsers += [pscustomobject] @{
                        'Policy Name'  = $Policy.Name
                        Name           = $User.Name
                        SamAccountName = $User.SamAccountName
                        Type           = $User.ObjectClass
                        SID            = $User.SID
                    }
                }
                foreach ($Group in $Groups) {
                    $PolicyUsers += [pscustomobject] @{
                        'Policy Name'  = $Policy.Name
                        Name           = $Group.Name
                        SamAccountName = $Group.SamAccountName
                        Type           = $Group.ObjectClass
                        SID            = $Group.SID
                    }
                }
            }
            #Get-AdFineGrainedPassowrdPolicySubject
            #Get-AdresultantPasswordPolicy -Identity <user>
            return $PolicyUsers
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainFineGrainedPoliciesUsersExtended)) {
        Write-Verbose "Getting domain information - $Domain DomainFineGrainedPoliciesUsersExtended"
        $Data.DomainFineGrainedPoliciesUsersExtended = Invoke-Command -ScriptBlock {
            $PolicyUsers = @()
            foreach ($Policy in $Data.DomainFineGrainedPolicies) {
                $Users = @()
                $Groups = @()
                foreach ($U in $Policy.'Applies To') {
                    $Users += Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $U
                    $Groups += Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainGroupsFullList -DistinguishedName $U
                }
                foreach ($User in $Users) {
                    $PolicyUsers += [pscustomobject] @{
                        'Policy Name'                       = $Policy.Name
                        Name                                = $User.Name
                        SamAccountName                      = $User.SamAccountName
                        Type                                = $User.ObjectClass
                        SID                                 = $User.SID
                        'High Privileged Group'             = 'N/A'
                        'Display Name'                      = $User.DisplayName
                        'Member Name'                       = $Member.Name
                        'User Principal Name'               = $User.UserPrincipalName
                        'Sam Account Name'                  = $User.SamAccountName
                        'Email Address'                     = $User.EmailAddress
                        'PasswordExpired'                   = $User.PasswordExpired
                        'PasswordLastSet'                   = $User.PasswordLastSet
                        'PasswordNotRequired'               = $User.PasswordNotRequired
                        'PasswordNeverExpires'              = $User.PasswordNeverExpires
                        'Enabled'                           = $User.Enabled
                        'MemberSID'                         = $Member.SID.Value
                        'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $User.Manager).Name
                        'ManagerEmail'                      = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $User.Manager).EmailAddress
                        'DateExpiry'                        = Convert-ToDateTime -Timestring $($Object."msDS-UserPasswordExpiryTimeComputed") # -Verbose
                        "DaysToExpire"                      = (Convert-TimeToDays -StartTime GET-DATE -EndTime (Convert-ToDateTime -Timestring $($User."msDS-UserPasswordExpiryTimeComputed")))
                        "AccountExpirationDate"             = $User.AccountExpirationDate
                        "AccountLockoutTime"                = $User.AccountLockoutTime
                        "AllowReversiblePasswordEncryption" = $User.AllowReversiblePasswordEncryption
                        "BadLogonCount"                     = $User.BadLogonCount
                        "CannotChangePassword"              = $User.CannotChangePassword
                        "CanonicalName"                     = $User.CanonicalName
                        'Given Name'                        = $User.GivenName
                        'Surname'                           = $User.Surname
                        "Description"                       = $User.Description
                        "DistinguishedName"                 = $User.DistinguishedName
                        "EmployeeID"                        = $User.EmployeeID
                        "EmployeeNumber"                    = $User.EmployeeNumber
                        "LastBadPasswordAttempt"            = $User.LastBadPasswordAttempt
                        "LastLogonDate"                     = $User.LastLogonDate
                        "Created"                           = $User.Created
                        "Modified"                          = $User.Modified
                        "Protected"                         = $User.ProtectedFromAccidentalDeletion
                        "Domain"                            = $Domain
                    }
                }

                foreach ($Group in $Groups) {
                    $GroupMembership = Get-ADGroupMember -Server $Domain -Identity $Group.SID -Recursive
                    foreach ($Member in $GroupMembership) {
                        $Object = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $Member.DistinguishedName)
                        $PolicyUsers += [pscustomobject][ordered] @{
                            'Policy Name'                       = $Policy.Name
                            Name                                = $Group.Name
                            SamAccountName                      = $Group.SamAccountName
                            Type                                = $Group.ObjectClass
                            SID                                 = $Group.SID
                            'High Privileged Group'             = if ($Group.adminCount -eq 1) { $True } else { $False }
                            'Display Name'                      = $Object.DisplayName
                            'Member Name'                       = $Member.Name
                            'User Principal Name'               = $Object.UserPrincipalName
                            'Sam Account Name'                  = $Object.SamAccountName
                            'Email Address'                     = $Object.EmailAddress
                            'PasswordExpired'                   = $Object.PasswordExpired
                            'PasswordLastSet'                   = $Object.PasswordLastSet
                            'PasswordNotRequired'               = $Object.PasswordNotRequired
                            'PasswordNeverExpires'              = $Object.PasswordNeverExpires
                            'Enabled'                           = $Object.Enabled
                            'MemberSID'                         = $Member.SID.Value
                            'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $Object.Manager).Name
                            'ManagerEmail'                      = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $Object.Manager).EmailAddress
                            'DateExpiry'                        = Convert-ToDateTime -Timestring $($Object."msDS-UserPasswordExpiryTimeComputed") # -Verbose
                            "DaysToExpire"                      = (Convert-TimeToDays -StartTime GET-DATE -EndTime (Convert-ToDateTime -Timestring $($Object."msDS-UserPasswordExpiryTimeComputed")))
                            "AccountExpirationDate"             = $Object.AccountExpirationDate
                            "AccountLockoutTime"                = $Object.AccountLockoutTime
                            "AllowReversiblePasswordEncryption" = $Object.AllowReversiblePasswordEncryption
                            "BadLogonCount"                     = $Object.BadLogonCount
                            "CannotChangePassword"              = $Object.CannotChangePassword
                            "CanonicalName"                     = $Object.CanonicalName
                            'Given Name'                        = $Object.GivenName
                            'Surname'                           = $Object.Surname
                            "Description"                       = $Object.Description
                            "DistinguishedName"                 = $Object.DistinguishedName
                            "EmployeeID"                        = $Object.EmployeeID
                            "EmployeeNumber"                    = $Object.EmployeeNumber
                            "LastBadPasswordAttempt"            = $Object.LastBadPasswordAttempt
                            "LastLogonDate"                     = $Object.LastLogonDate
                            "Created"                           = $Object.Created
                            "Modified"                          = $Object.Modified
                            "Protected"                         = $Object.ProtectedFromAccidentalDeletion
                            "Domain"                            = $Domain
                        }
                    }
                }


            }
            return $PolicyUsers
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainGroups, [ActiveDirectory]::DomainGroupsSpecial)) {
        Write-Verbose "Getting domain information - $Domain DomainGroups"
        $Data.DomainGroups = Get-WinGroups -Groups $Data.DomainGroupsFullList -Users $Data.DomainUsersFullList -Domain $Domain
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainGroups, [ActiveDirectory]::DomainGroupsMembers)) {
        Write-Verbose "Getting domain information - $Domain DomainGroupsMembers"
        $Data.DomainGroupsMembers = Get-WinGroupMembers -Groups $Data.DomainGroups -Domain $Domain -ADCatalog $Data.DomainUsersFullList, $Data.DomainComputersFullList, $Data.DomainGroupsFullList -ADCatalogUsers $Data.DomainUsersFullList -Option Standard
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainGroups, [ActiveDirectory]::DomainGroupsMembersRecursive)) {
        Write-Verbose "Getting domain information - $Domain DomainGroupsMembersRecursive"
        $Data.DomainGroupsMembersRecursive = Get-WinGroupMembers -Groups $Data.DomainGroups -Domain $Domain -ADCatalog $Data.DomainUsersFullList, $Data.DomainComputersFullList, $Data.DomainGroupsFullList -ADCatalogUsers $Data.DomainUsersFullList -Option Recursive
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainGroupsPriviliged, [ActiveDirectory]::DomainGroupMembersRecursivePriviliged)) {
        Write-Verbose "Getting domain information - $Domain DomainGroupsPriviliged"
        $PrivilegedGroupsSID = "S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549", "S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552", "S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573", "S-1-5-32-578", "S-1-5-32-580", "$($Data.DomainInformation.DomainSID)-512", "$($Data.DomainInformation.DomainSID)-518", "$($Data.DomainInformation.DomainSID)D-519", "$($Data.DomainInformation.DomainSID)-520"
        $Data.DomainGroupsPriviliged = $Data.DomainGroups | Where { $PrivilegedGroupsSID -contains $_.'Group SID' }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainGroupsSpecial, [ActiveDirectory]::DomainGroupMembersRecursiveSpecial)) {
        Write-Verbose "Getting domain information - $Domain DomainGroupsSpecial"
        $Data.DomainGroupsSpecial = $Data.DomainGroups | Where { ($_.'Group SID').Length -eq 12 }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainGroupsSpecialMembers, [ActiveDirectory]::DomainGroupsSpecialMembersRecursive)) {
        Write-Verbose "Getting domain information - $Domain DomainGroupMembersSpecialRecursive"
        $Data.DomainGroupsSpecialMembers = $Data.DomainGroupsMembers  | Where { ($_.'Group SID').Length -eq 12 } | Select-Object * #-Exclude Group*, 'High Privileged Group'
        Write-Verbose "Getting domain information - $Domain DomainGroupsSpecialMembersRecursive"
        $Data.DomainGroupsSpecialMembersRecursive = $Data.DomainGroupsMembersRecursive  | Where { ($_.'Group SID').Length -eq 12 } | Select-Object * #-Exclude Group*, 'High Privileged Group'
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainGroupsPriviligedMembers, [ActiveDirectory]::DomainGroupsPriviligedMembersRecursive)) {
        Write-Verbose "Getting domain information - $Domain DomainGroupsPriviligedMembers"
        $Data.DomainGroupsPriviligedMembers = $Data.DomainGroupsMembers  | Where { $Data.DomainGroupsPriviliged.'Group SID' -contains ($_.'Group SID') } | Select-Object * #-Exclude Group*, 'High Privileged Group'
        Write-Verbose "Getting domain information - $Domain DomainGroupsPriviligedMembersRecursive"
        $Data.DomainGroupsPriviligedMembersRecursive = $Data.DomainGroupsMembersRecursive  | Where { $Data.DomainGroupsPriviliged.'Group SID' -contains ($_.'Group SID') } | Select-Object * #-Exclude Group*, 'High Privileged Group'
    }
    ## Users per one group only.
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainAdministrators, [ActiveDirectory]::DomainGroupsMembers)) {
        Write-Verbose "Getting domain information - $Domain DomainAdministrators"
        $Data.DomainAdministrators = $Data.DomainGroupsMembers  | Where { $_.'Group SID' -eq $('{0}-512' -f $Data.DomainInformation.DomainSID.Value) } | Select-Object * -Exclude Group*, 'High Privileged Group'
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainAdministratorsRecursive, [ActiveDirectory]::DomainGroupsMembersRecursive)) {
        Write-Verbose "Getting domain information - $Domain DomainAdministratorsRecursive"
        $Data.DomainAdministratorsRecursive = $Data.DomainGroupsMembersRecursive  | Where { $_.'Group SID' -eq $('{0}-512' -f $Data.DomainInformation.DomainSID.Value) } | Select-Object * -Exclude Group*, 'High Privileged Group'
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainEnterpriseAdministrators, [ActiveDirectory]::DomainGroupsMembers)) {
        Write-Verbose "Getting domain information - $Domain DomainEnterpriseAdministrators"
        $Data.DomainEnterpriseAdministrators = $Data.DomainGroupsMembers | Where { $_.'Group SID' -eq $('{0}-519' -f $Data.DomainInformation.DomainSID.Value) } | Select-Object * -Exclude Group*, 'High Privileged Group'
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainEnterpriseAdministratorsRecursive, [ActiveDirectory]::DomainGroupsMembersRecursive)) {
        Write-Verbose "Getting domain information - $Domain DomainEnterpriseAdministratorsRecursive"
        $Data.DomainEnterpriseAdministratorsRecursive = $Data.DomainGroupsMembersRecursive | Where { $_.'Group SID' -eq $('{0}-519' -f $Data.DomainInformation.DomainSID.Value) } | Select-Object * -Exclude Group*, 'High Privileged Group'
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [ActiveDirectory]::DomainPasswordDataUsers,
            [ActiveDirectory]::DomainPasswordDataPasswords,
            [ActiveDirectory]::DomainPasswordClearTextPassword,
            [ActiveDirectory]::DomainPasswordLMHash,
            [ActiveDirectory]::DomainPasswordEmptyPassword,
            [ActiveDirectory]::DomainPasswordWeakPassword,
            [ActiveDirectory]::DomainPasswordDefaultComputerPassword,
            [ActiveDirectory]::DomainPasswordPasswordNotRequired,
            [ActiveDirectory]::DomainPasswordPasswordNeverExpires,
            [ActiveDirectory]::DomainPasswordAESKeysMissing,
            [ActiveDirectory]::DomainPasswordPreAuthNotRequired,
            [ActiveDirectory]::DomainPasswordDESEncryptionOnly,
            [ActiveDirectory]::DomainPasswordDelegatableAdmins,
            [ActiveDirectory]::DomainPasswordDuplicatePasswordGroups,
            [ActiveDirectory]::DomainPasswordStats,
            [ActiveDirectory]::DomainPasswordHashesWeakPassword
        )) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDataUsers - This will take a while if set!"
        $TimeToProcess = Start-TimeLog
        $Data.DomainPasswordDataUsers = Get-ADReplAccount -All -Server $Data.DomainInformation.DnsRoot -NamingContext $Data.DomainInformation.DistinguishedName
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDataUsers - Time: $($TimeToProcess | Stop-TimeLog)"
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [ActiveDirectory]::DomainPasswordDataPasswords,
            [ActiveDirectory]::DomainPasswordClearTextPassword,
            [ActiveDirectory]::DomainPasswordLMHash,
            [ActiveDirectory]::DomainPasswordEmptyPassword,
            [ActiveDirectory]::DomainPasswordWeakPassword,
            [ActiveDirectory]::DomainPasswordDefaultComputerPassword,
            [ActiveDirectory]::DomainPasswordPasswordNotRequired,
            [ActiveDirectory]::DomainPasswordPasswordNeverExpires,
            [ActiveDirectory]::DomainPasswordAESKeysMissing,
            [ActiveDirectory]::DomainPasswordPreAuthNotRequired,
            [ActiveDirectory]::DomainPasswordDESEncryptionOnly,
            [ActiveDirectory]::DomainPasswordDelegatableAdmins,
            [ActiveDirectory]::DomainPasswordDuplicatePasswordGroups,
            [ActiveDirectory]::DomainPasswordStats,
            [ActiveDirectory]::DomainPasswordHashesWeakPassword
        )) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDataPasswords - This will take a while if set!"
        $TimeToProcess = Start-TimeLog
        $Data.DomainPasswordDataPasswords = Get-WinADDomainPasswordQuality -FilePath $PathToPasswords -DomainInformation $Data -Verbose:$false -PasswordQualityUsers $Data.DomainPasswordDataUsers
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDataPasswords - Time: $($TimeToProcess | Stop-TimeLog)"
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainPasswordHashesWeakPassword)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDataPasswordsHashes - This will take a while if set!"
        $TimeToProcess = Start-TimeLog
        $Data.DomainPasswordDataPasswordsHashes = Get-WinADDomainPasswordQuality -FilePath $PathToPasswordsHashes -DomainInformation $Data -UseHashes -Verbose:$false -PasswordQualityUsers $Data.DomainPasswordDataUsers
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDataPasswordsHashes - Time: $($TimeToProcess | Stop-TimeLog)"
    }
    if ($Data.DomainPasswordDataPasswords) {
        $PasswordsQuality = $Data.DomainPasswordDataPasswords
    } elseif ($Data.DomainPasswordDataPasswordsHashes) {
        $PasswordsQuality = $Data.DomainPasswordDataPasswordsHashes
    } else {
        $PasswordsQuality = $null
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordClearTextPassword)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordClearTextPassword"
        $Data.DomainPasswordClearTextPassword = $PasswordsQuality.DomainPasswordClearTextPassword
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordLMHash)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordLMHash"
        $Data.DomainPasswordLMHash = $PasswordsQuality.DomainPasswordLMHash
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordEmptyPassword)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordEmptyPassword"
        $Data.DomainPasswordEmptyPassword = $PasswordsQuality.DomainPasswordEmptyPassword
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordWeakPassword)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordWeakPassword"
        $Data.DomainPasswordWeakPassword = $PasswordsQuality.DomainPasswordWeakPassword
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordDefaultComputerPassword)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDefaultComputerPassword"
        $Data.DomainPasswordDefaultComputerPassword = $PasswordsQuality.DomainPasswordDefaultComputerPassword
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordPasswordNotRequired)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordPasswordNotRequired"
        $Data.DomainPasswordPasswordNotRequired = $PasswordsQuality.DomainPasswordPasswordNotRequired
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordPasswordNeverExpires)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordPasswordNeverExpires"
        $Data.DomainPasswordPasswordNeverExpires = $PasswordsQuality.DomainPasswordPasswordNeverExpires
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordAESKeysMissing)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordAESKeysMissing"
        $Data.DomainPasswordAESKeysMissing = $PasswordsQuality.DomainPasswordAESKeysMissing
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordPreAuthNotRequired)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordPreAuthNotRequired"
        $Data.DomainPasswordPreAuthNotRequired = $PasswordsQuality.DomainPasswordPreAuthNotRequired
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordDESEncryptionOnly)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDESEncryptionOnly"
        $Data.DomainPasswordDESEncryptionOnly = $PasswordsQuality.DomainPasswordDESEncryptionOnly
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordDelegatableAdmins)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDelegatableAdmins"
        $Data.DomainPasswordDelegatableAdmins = $PasswordsQuality.DomainPasswordDelegatableAdmins
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordDuplicatePasswordGroups)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordDuplicatePasswordGroups"
        $Data.DomainPasswordDuplicatePasswordGroups = $PasswordsQuality.DomainPasswordDuplicatePasswordGroups
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordHashesWeakPassword)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordHashesWeakPassword"
        $Data.DomainPasswordHashesWeakPassword = $Data.DomainPasswordDataPasswordsHashes.DomainPasswordWeakPassword
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @( [ActiveDirectory]::DomainPasswordStats)) {
        Write-Verbose "Getting domain password information - $Domain DomainPasswordStats"
        $Data.DomainPasswordStats = Invoke-Command -ScriptBlock {
            $Stats = [ordered] @{}
            $Stats.'Clear Text Passwords' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordClearTextPassword
            $Stats.'LM Hashes' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordLMHash
            $Stats.'Empty Passwords' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordEmptyPassword
            $Stats.'Weak Passwords' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordWeakPassword
            if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainPasswordHashesWeakPassword)) {
                $Stats.'Weak Passwords (HASH)' = Get-ObjectCount -Object $Data.DomainPasswordDataPasswordsHashes.DomainPasswordHashesWeakPassword
            }
            $Stats.'Default Computer Passwords' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordDefaultComputerPassword
            $Stats.'Password Not Required' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordPasswordNotRequired
            $Stats.'Password Never Expires' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordPasswordNeverExpires
            $Stats.'AES Keys Missing' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordAESKeysMissing
            $Stats.'PreAuth Not Required' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordPreAuthNotRequired
            $Stats.'DES Encryption Only' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordDESEncryptionOnly
            $Stats.'Delegatable Admins' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordDelegatableAdmins
            $Stats.'Duplicate Password Users' = Get-ObjectCount -Object $PasswordsQuality.DomainPasswordDuplicatePasswordGroups
            $Stats.'Duplicate Password Grouped' = Get-ObjectCount ($PasswordsQuality.DomainPasswordDuplicatePasswordGroups.'Duplicate Group' | Sort-Object -Unique)
            return $Stats
        }
    }
    return $Data
}
