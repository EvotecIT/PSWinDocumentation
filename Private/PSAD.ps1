#Import-Module ActiveDirectory
function Get-ActiveDirectoryData {
    param(

    )
    $ADSnapshot = @{}
    $ADSnapshot.RootDSE = $(
        $Info = Get-ADRootDSE
        <#
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
        #$Info
        $Info | Select-Object `
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

    )
    $ADSnapshot.FSMO = $( #Flexible Single Master Operations
        #$Info1 = Get-ADRootDSE | Select-Object DomainNamingMaster
        $Info2 = Get-ADDomain | Select-Object PDCEmulator, RIDMaster, InfrastructureMaster
        $Info3 = Get-ADForest | Select-Object SchemaMaster, DomainNamingMaster
        $Info = @{
            'Domain Naming Master'  = $Info3.DomainNamingMaster
            'Schema Master'         = $Info3.SchemaMaster
            'PDC Emulator'          = $Info2.PDCEmulator
            'RID Master'            = $Info2.RIDMaster
            'Infrastructure Master' = $Info2.InfrastructureMaster
        }
        $Info
    )
    $ADSnapshot.ForestInformation = $(
        $Info = Get-ADForest
        <#
            ApplicationPartitions
            CrossForestReferences
            DomainNamingMaster
            Domains
            ForestMode
            GlobalCatalogs
            Name
            PartitionsContainer
            RootDomain
            SchemaMaster
            Sites
            SPNSuffixes
            UPNSuffixes
        #>

        $Info | Select-Object Name, `
        @{label = 'Forest Mode'; expression = { $_.ForestMode }},
        @{label = 'Domain Naming Master'; expression = { $_.DomainNamingMaster }},
        Domains,
        Sites,
        SPNSuffixes,
        UPNSuffixes,
        SchemaMaster,
        RootDomain
    )
    $ADSnapshot.ForestFeatures = $(
        $Info = @{ 'Recycle Bin Enabled' = "$(Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"').EnabledScopes)" -eq '' }
        $Info
    )


    $ADSnapshot.DomainInformation = $(
        Get-ADDomain
    )
    $ADSnapshot.DomainControllers = $(Get-ADDomainController -Filter *)
    $ADSnapshot.DomainTrusts = (Get-ADTrust -Filter *)
    $ADSnapshot.DefaultPassWordPoLicy = $(Get-ADDefaultDomainPasswordPolicy)
    $ADSnapshot.AuthenticationPolicies = $(Get-ADAuthenticationPolicy -LDAPFilter '(name=AuthenticationPolicy*)')
    $ADSnapshot.AuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Filter 'Name -like "*AuthenticationPolicySilo*"')
    $ADSnapshot.CentralAccessPolicies = $(Get-ADCentralAccessPolicy -Filter *)
    $ADSnapshot.CentralAccessRules = $(Get-ADCentralAccessRule -Filter *)
    $ADSnapshot.ClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Filter *)
    $ADSnapshot.ClaimTypes = $(Get-ADClaimType -Filter *)
    $ADSnapshot.DomainAdministrators = $( Get-ADGroup -Identity $('{0}-512' -f (Get-ADDomain).domainSID) | Get-ADGroupMember -Recursive)
    $ADSnapshot.OrganizationalUnits = $(Get-ADOrganizationalUnit -Filter *)
    $ADSnapshot.OptionalFeatures = $(Get-ADOptionalFeature -Filter *)
    $ADSnapshot.Sites = $(Get-ADReplicationSite -Filter *)
    $ADSnapshot.Subnets = $(Get-ADReplicationSubnet -Filter *)
    $ADSnapshot.SiteLinks = $(Get-ADReplicationSiteLink -Filter *)
    $ADSnapshot.LDAPDNS = $(Resolve-DnsName -Name "_ldap._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.KerberosDNS = $(Resolve-DnsName -Name "_kerberos._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.GroupPolicies = $(
        $Info = Get-GPO -All | Select DisplayName, Owner, DomainName, CreationTime, ModificationTime, GpoStatus, WmiFilter, Description # Id, UserVersion, ComputerVersion
        $Info | Select-Object DisplayName, CreationTime, ModificationTime, GpoStatus, WmiFilter, Description
    )
    return $ADSnapshot
}
Clear-Host
$ADSnapshot = Get-ActiveDirectoryData
Write-Color 'Forest Information' -Color Red
$ADSnapshot.ForestInformation
Write-Color 'RootDSE' -Color Red
$ADSnapshot.RootDSE
Write-Color 'Domain Information' -Color Red
$ADSnapshot.DomainInformation
Write-Color 'FSMO Information' -Color Red
$ADSnapshot.FSMO
Write-Color 'Forest Features' -Color Red
$ADSnapshot.ForestFeatures | ft -a
#$ADSnapshot.DomainInformation
#$ADSnapshot.GroupPolicies | ft -AutoSize