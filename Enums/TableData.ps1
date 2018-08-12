Add-Type -TypeDefinition @"
    public enum Forest {
        ForestInformation,
        FSMO,
        GlobalCatalogs,
        OptionalFeatures,
        UPNSuffixes,
        SPNSuffixes,
        Sites,
        Sites1,
        Sites2,
        Subnets,
        Subnets1,
        Subnets2,
        SiteLinks
    }
"@

Add-Type -TypeDefinition @"
    public enum Domain {
        AuthenticationPolicies, // Not yet tested
        AuthenticationPolicySilos, // Not yet tested
        CentralAccessPolicies, // Not yet tested
        CentralAccessRules, // Not yet tested
        ClaimTransformPolicies, // Not yet tested
        ClaimTypes, // Not yet tested
        LDAPDNS, // not yet tested
        KerberosDNS, // not yet tested
        DomainAdministrators,
        DomainInformation,
        DomainControllers,
        FSMO,
        DefaultPasswordPoLicy,
        GroupPolicies,
        GroupPoliciesDetails,
        GroupPoliciesACL,
        OrganizationalUnits,
        PriviligedGroupMembers,
        Users,
        UsersCount,
        DomainTrusts
    }
"@