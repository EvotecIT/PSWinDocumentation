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
        DomainControllers,
        FSMO,
        DefaultPasswordPoLicy,
        GroupPolicies,
        OrganizationalUnits,
        PriviligedGroupMembers,
        DomainAdministrators,
        UsersCount
    }
"@