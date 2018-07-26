#Import-Module ActiveDirectory
function Get-ActiveDirectoryCleanData {
    param()
    $ADSnapshot = @{}
    <# RootDSE
                configurationNamingContext
                currentTime
                defaultNamingContext
                dnsHostName
                domainControllerFunctionality
                domainFunctionality
                dsServiceName
                forestFunctionality
                highestCommittedUSN
                isGlobalCatalogReady
                isSynchronized
                ldapServiceName
                namingContexts
                rootDomainNamingContext
                schemaNamingContext
                serverName
                subschemaSubentry
                supportedCapabilities
                supportedControl
                supportedLDAPPolicies
                supportedLDAPVersion
                supportedSASLMechanisms
    #>
    $ADSnapshot.RootDSE = $(Get-ADRootDSE)
    $ADSnapshot.ForestInformation = $(Get-ADForest)
    $ADSnapshot.DomainInformation = $(Get-ADDomain)
    $ADSnapshot.DomainControllers = $(Get-ADDomainController -Filter * )
    $ADSnapshot.DomainTrusts = (Get-ADTrust -Filter * )
    $ADSnapshot.DefaultPassWordPoLicy = $(Get-ADDefaultDomainPasswordPolicy)
    $ADSnapshot.AuthenticationPolicies = $(Get-ADAuthenticationPolicy -LDAPFilter '(name=AuthenticationPolicy*)')
    $ADSnapshot.AuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Filter 'Name -like "*AuthenticationPolicySilo*"')
    $ADSnapshot.CentralAccessPolicies = $(Get-ADCentralAccessPolicy -Filter * )
    $ADSnapshot.CentralAccessRules = $(Get-ADCentralAccessRule -Filter * )
    $ADSnapshot.ClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Filter * )
    $ADSnapshot.ClaimTypes = $(Get-ADClaimType -Filter * )
    $ADSnapshot.DomainAdministrators = $( Get-ADGroup -Identity $('{0}-512' -f (Get-ADDomain).domainSID) | Get-ADGroupMember -Recursive)
    $ADSnapshot.OrganizationalUnits = $(Get-ADOrganizationalUnit -Filter * )
    $ADSnapshot.OptionalFeatures = $(Get-ADOptionalFeature -Filter * )
    $ADSnapshot.Sites = $(Get-ADReplicationSite -Filter * )
    $ADSnapshot.Subnets = $(Get-ADReplicationSubnet -Filter * )
    $ADSnapshot.SiteLinks = $(Get-ADReplicationSiteLink -Filter * )
    $ADSnapshot.LDAPDNS = $(Resolve-DnsName -Name "_ldap._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.KerberosDNS = $(Resolve-DnsName -Name "_kerberos._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.GroupPolicies = $(Get-GPO -All) # DisplayName, Owner, DomainName, CreationTime, ModificationTime, GpoStatus, WmiFilter, Description # Id, UserVersion, ComputerVersion
    return $ADSnapshot
}

function Get-ActiveDirectoryProcessedData {
    $ADSnapshot = Get-ActiveDirectoryCleanData

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
    $DisplayAD.GroupPolicies = [ordered] @{
        'Display Name'      = $ADSnapshot.GroupPolicies.DisplayName
        'Creation Time'     = $ADSnapshot.GroupPolicies.CreationTime
        'Modification Time' = $ADSnapshot.GroupPolicies.ModificationTime
        'Gpo Status'        = $ADSnapshot.GroupPolicies.GPOStatus
        'Wmi Filter'        = $ADSnapshot.GroupPolicies.WmiFilter
        'Description'       = $ADSnapshot.GroupPolicies.Description
    }
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


    return $DisplayAD
}
<#
$ADSnapshot = Get-ActiveDirectoryCleanData
$AD = Get-ActiveDirectoryProcessedData

$ADSnapshot.DomainInformation

$ADSnapshot.RootDSE

$ADSnapshot.ForestInformation
#>


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