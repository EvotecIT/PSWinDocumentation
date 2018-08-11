Add-Type -TypeDefinition @"
    public enum Forest {
        Summary,
        FSMO,
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
        AuthenticationPolicies,
        AuthenticationPolicySilos,
        CentralAccessPolicies,
        CentralAccessRules,
        ClaimTransformPolicies,
        ClaimTypes, // Not yet done
        LDAPDNS, // not yet done
        KerberosDNS, // not yet done
        DomainAdministrators,
        DomainInformation,
        DomainControllers,
        FSMO,
        DefaultPasswordPoLicy,
        GroupPolicies,
        GroupPoliciesDetails,
        OrganizationalUnits,
        PriviligedGroupMembers,
        Users,
        UsersCount,
        DomainTrusts
    }
"@