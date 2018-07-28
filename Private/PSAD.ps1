function Get-ActiveDirectoryCleanData {
    param(
        $Domain
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
    $ADSnapshot.DomainAdministrators = $( Get-ADGroup -Identity $('{0}-512' -f (Get-ADDomain).domainSID) | Get-ADGroupMember -Recursive)
    $ADSnapshot.OrganizationalUnits = $(Get-ADOrganizationalUnit -Server $Domain -Filter * )
    $ADSnapshot.OptionalFeatures = $(Get-ADOptionalFeature -Server $Domain -Filter * )
    $ADSnapshot.Sites = $(Get-ADReplicationSite -Server $Domain -Filter * )
    $ADSnapshot.Subnets = $(Get-ADReplicationSubnet -Server $Domain -Filter * )
    $ADSnapshot.SiteLinks = $(Get-ADReplicationSiteLink -Server $Domain -Filter * )
    $ADSnapshot.LDAPDNS = $(Resolve-DnsName -Name "_ldap._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.KerberosDNS = $(Resolve-DnsName -Name "_kerberos._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.GroupPolicies = $(Get-GPO -Domain $Domain -All) # DisplayName, Owner, DomainName, CreationTime, ModificationTime, GpoStatus, WmiFilter, Description # Id, UserVersion, ComputerVersion
    return $ADSnapshot
}

function Get-ActiveDirectoryProcessedData {
    param (
        $ADSnapshot
    )
    $DisplayAD = @{}
    $DisplayAD.RootDSE = $ADSnapshot.RootDSE
    $DisplayAD.DomainInformation = $ADSnapshot.DomainInformation
    $DisplayAD.FSMO = [ordered] @{
        'Domain Naming Master'  = $ADSnapshot.ForestInformation.DomainNamingMaster
        'Schema Master'         = $ADSnapshot.ForestInformation.SchemaMaster
        'PDC Emulator'          = $ADSnapshot.DomainInformation.PDCEmulator
        'RID Master'            = $ADSnapshot.DomainInformation.RIDMaster
        'Infrastructure Master' = $ADSnapshot.DomainInformation.InfrastructureMaster
    }

    $GroupPolicies = @()
    foreach ($gpo in $ADSnapshot.GroupPolicies) {

        $GroupPolicy = [ordered] @{
            'Display Name'      = $gpo.DisplayName
            'Creation Time'     = $gpo.CreationTime
            'Modification Time' = $gpo.ModificationTime
            'Gpo Status'        = $gpo.GPOStatus
            'Wmi Filter'        = $gpo.WmiFilter
            'Description'       = $gpo.Description
        }
        $GroupPolicies += $GroupPolicy
    }
    $DisplayAD.GroupPolicies = $GroupPolicies
    $DisplayAD.GroupPoliciesTable = $GroupPolicies.ForEach( {[PSCustomObject]$_})
    $DisplayAD.ForestInformation = [ordered] @{
        'Name'                    = $ADSnapshot.ForestInformation.Name
        'Root Domain'             = $ADSnapshot.ForestInformation.RootDomain
        'Forest Functional Level' = $ADSnapshot.ForestInformation.ForestMode
        'Domains Count'           = ($ADSnapshot.ForestInformation.Domains).Count
        'Sites Count'             = ($ADSnapshot.ForestInformation.Sites).Count
        'Domains'                 = ($ADSnapshot.ForestInformation.Domains) -join ", "
        'Sites'                   = ($ADSnapshot.ForestInformation.Sites) -join ", "
    }
    $DisplayAD.OptionalFeatures = [ordered] @{
        'Recycle Bin Enabled'                          = ''
        #'Recycle Bin Scopes' = ''
        'Privileged Access Management Feature Enabled' = ''
        #'Privileged Access Management Feature Scopes' ''
    }
    ### Fix Optional Features
    foreach ($Feature in $ADSnapshot.OptionalFeatures) {
        if ($Feature.Name -eq 'Recycle Bin Feature') {
            if ("$($Feature.EnabledScopes)" -eq '') { $DisplayAD.OptionalFeatures.'Recycle Bin Enabled' = $False }
            else { $DisplayAD.OptionalFeatures.'Recycle Bin Enabled' = $True }
        }
        if ($Feature.Name -eq 'Privileged Access Management Feature') {
            if ("$($Feature.EnabledScopes)" -eq '') { $DisplayAD.OptionalFeatures.'Privileged Access Management Feature Enabled' = $False }
            else { $DisplayAD.OptionalFeatures.'Privileged Access Management Feature Enabled' = $True }
        }
    }
    ### Fix optional features
    $UPNSuffixList = @()
    $UPNSuffixList += $ADSnapshot.ForestInformation.RootDomain + ' (Primary/Default UPN)'
    $UPNSuffixList += $ADSnapshot.ForestInformation.UPNSuffixes
    $DisplayAD.UPNSuffixes = $UPNSuffixList


    $DisplayAD.DefaultPassWordPoLicy = [ordered] @{
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

    return $DisplayAD
}

#Get-ADRootDSE -Server 'ad.evotec.xyz'
#Write-Color 'Next' -Color Red
#Get-AdRootDSe -Server 'ad.evotec.pl'
#Write-Color 'Next' -Color Red
#Get-AdForest -Server 'ad.evotec.xyz'
#Write-Color 'Next' -Color Red
#Get-AdForest -Server 'ad.evotec.pl'


#$ADSnapshot.DomainInformation

#$ADSnapshot.RootDSE

#$ADSnapshot.ForestInformation

#$ADSnapshot.DefaultPassWordPoLicy
#Get-ADDefaultDomainPasswordPolicy -Server 'ad.evotec.xyz'
#Get-ADDefaultDomainPasswordPolicy -Server 'ad.evotec.pl'


#Get-ADDomain -Server 'ad.evotec.xyz'
#Et-AdDomain -Server 'ad.evotec.pl'
#( + $ADSnapshot.ForestInformation.UPNSuffixes) -join ', '

#foreach ($Element in $($ADSnapshot.OptionalFeatures).PSObject.Propet) {
#    $Element
#} #.EnabledScopes

<#
$Info = $ADSnapshot.RootDSE | Select-Object `
@{label = 'Configuration Naming Context'; expression = { $_.configurationNamingContext }},
@{label = 'Default Naming Context'; expression = { $_.defaultNamingContext }},
@{label = 'DNS Host Name'; expression = { $_.dnsHostName }},
@{label = 'Domain Controller Functionality'; expression = { $_.domainControllerFunctionality }},
@{label = 'Domain Functionality'; expression = { $_.domainFunctionality }},
@{label = 'Forest Functionality'; expression = { $_.forestFunctionality }},
@{label = 'Supported LDAP Policies'; expression = { $_.supportedLDAPPolicies }},
@{label = 'Sub Schema Subentry'; expression = { $_.subschemaSubentry }},
@{label = 'Supported LDAP Version'; expression = { $_.supportedLDAPVersion }},
@{label = 'Supported SASL Mechanisms'; expression = { $_.supportedSASLMechanisms }}

$Info1 = Get-ADForest | Select-Object ApplicationPartitions,
CrossForestReferences,
DomainNamingMaster,
Domains,
ForestMode,
GlobalCatalogs,
Name,
PartitionsContainer,
RootDomain,
SchemaMaster,
Sites,
SPNSuffixes,
UPNSuffixes,


#>