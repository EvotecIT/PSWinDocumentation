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
        defaultNamingContext,
        dnsHostName,
        domainControllerFunctionality,
        domainFunctionality,
        forestFunctionality,
        supportedLDAPPolicies,
        subschemaSubentry,
        supportedLDAPVersion,
        supportedSASLMechanisms

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

        $Info | Select-Object Name, ForestMode, DomainNamingMaster, Domains, Sites
    )
    $ADSnapshot.DomainInformation = $(Get-ADDomain)
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
    return $ADSnapshot
}