function Get-ActiveDirectoryData {
    param(
        [string] $Domain
    )
    $ADSnapshot = @{}
    $ADSnapshot.RootDSE = $(Get-ADRootDSE -Server $Domain)
    $ADSnapshot.ForestInformation = $(Get-ADForest -Server $Domain)
    $ADSnapshot.DomainInformation = $(Get-ADDomain -Server $Domain)
    $ADSnapshot.DomainControllers = $(Get-ADDomainController -Server $Domain -Filter * )
    $ADSnapshot.DomainTrusts = (Get-ADTrust -Server $Domain -Filter * )
    $ADSnapshot.DefaultPassWordPoLicy = $(Get-ADDefaultDomainPasswordPolicy -Server $Domain)
    $ADSnapshot.AuthenticationPolicies = $(Get-ADAuthenticationPolicy -Server $Domain -LDAPFilter '(name=AuthenticationPolicy*)')
    $ADSnapshot.AuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Server $Domain -Filter 'Name -like "*AuthenticationPolicySilo*"')
    $ADSnapshot.CentralAccessPolicies = $(Get-ADCentralAccessPolicy -Server $Domain -Filter * )
    $ADSnapshot.CentralAccessRules = $(Get-ADCentralAccessRule -Server $Domain -Filter * )
    $ADSnapshot.ClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Server $Domain -Filter * )
    $ADSnapshot.ClaimTypes = $(Get-ADClaimType -Server $Domain -Filter * )
    $ADSnapshot.DomainAdministrators = $( Get-ADGroup -Identity $('{0}-512' -f (Get-ADDomain).domainSID) | Get-ADGroupMember -Recursive | Get-ADUser)
    $ADSnapshot.OrganizationalUnits = $(Get-ADOrganizationalUnit -Server $Domain -Properties * -Filter * )
    $ADSnapshot.Sites = $(Get-ADReplicationSite -Server $Domain -Filter * )
    $ADSnapshot.Subnets = $(Get-ADReplicationSubnet -Server $Domain -Filter * )
    $ADSnapshot.SiteLinks = $(Get-ADReplicationSiteLink -Server $Domain -Filter * )
    $ADSnapshot.LDAPDNS = $(Resolve-DnsName -Name "_ldap._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.KerberosDNS = $(Resolve-DnsName -Name "_kerberos._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.GroupPolicies = $(Get-GPO -Domain $Domain -All) # DisplayName, Owner, DomainName, CreationTime, ModificationTime, GpoStatus, WmiFilter, Description # Id, UserVersion, ComputerVersion
    return $ADSnapshot
}

function Get-WinDomainInformation {
    param (
        [string] $Domain
    )
    $ADSnapshot = Get-ActiveDirectoryData -Domain $Domain
    $Data = @{}
    $Data.ADSnapshot = $ADSnapshot
    $Data.RootDSE = $ADSnapshot.RootDSE
    $Data.DomainInformation = $ADSnapshot.DomainInformation
    $Data.FSMO = [ordered] @{
        #'Domain Naming Master'  = $ADSnapshot.ForestInformation.DomainNamingMaster
        #'Schema Master'         = $ADSnapshot.ForestInformation.SchemaMaster
        'PDC Emulator'          = $ADSnapshot.DomainInformation.PDCEmulator
        'RID Master'            = $ADSnapshot.DomainInformation.RIDMaster
        'Infrastructure Master' = $ADSnapshot.DomainInformation.InfrastructureMaster
    }
    $Data.GroupPolicies = Invoke-Command -ScriptBlock {
        $GroupPolicies = @()
        foreach ($gpo in $ADSnapshot.GroupPolicies) {
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
        return $GroupPolicies.ForEach( {[PSCustomObject]$_})
    }
    $Data.ForestInformation = [ordered] @{
        'Name'                    = $ADSnapshot.ForestInformation.Name
        'Root Domain'             = $ADSnapshot.ForestInformation.RootDomain
        'Forest Functional Level' = $ADSnapshot.ForestInformation.ForestMode
        'Domains Count'           = ($ADSnapshot.ForestInformation.Domains).Count
        'Sites Count'             = ($ADSnapshot.ForestInformation.Sites).Count
        'Domains'                 = ($ADSnapshot.ForestInformation.Domains) -join ", "
        'Sites'                   = ($ADSnapshot.ForestInformation.Sites) -join ", "
    }
    $Data.DefaultPassWordPoLicy = [ordered] @{
        'Complexity Enabled'            = $ADSnapshot.DefaultPassWordPoLicy.ComplexityEnabled
        #'Distinguished Name'            = $ADSnapshot.DefaultPassWordPoLicy.DistinguishedName
        'Lockout Duration'              = $ADSnapshot.DefaultPassWordPoLicy.LockoutDuration
        'Lockout Observation Window'    = $ADSnapshot.DefaultPassWordPoLicy.LockoutObservationWindow
        'Lockout Threshold'             = $ADSnapshot.DefaultPassWordPoLicy.LockoutThreshold
        'Max Password Age'              = $ADSnapshot.DefaultPassWordPoLicy.MaxPasswordAge
        'Min Password Age'              = $ADSnapshot.DefaultPassWordPoLicy.MinPasswordAge
        'Min Password Length'           = $ADSnapshot.DefaultPassWordPoLicy.MinPasswordAge
        'Password History Count'        = $ADSnapshot.DefaultPassWordPoLicy.PasswordHistoryCount
        'Reversible Encryption Enabled' = $ADSnapshot.DefaultPassWordPoLicy.ReversibleEncryptionEnabled
    }
    $Data.PriviligedGroupMembers = Get-PrivilegedGroupsMembers -Domain $Data.DomainInformation.DNSRoot -DomainSID $Data.DomainInformation.DomainSid
    $Data.OrganizationalUnits = $ADSnapshot.OrganizationalUnits | Select-Object Name, CanonicalName, Created | Sort-Object CanonicalName
    $Data.DomainAdministrators = $ADSnapshot.DomainAdministrators | Select-Object Name, SamAccountName, UserPrincipalName, Enabled
    return $Data
}