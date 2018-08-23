Add-Type -TypeDefinition @"
    public enum ActiveDirectory {
        ForestInformation,
        ForestFSMO,
        ForestGlobalCatalogs,
        ForestOptionalFeatures,
        ForestUPNSuffixes,
        ForestSPNSuffixes,
        ForestSites,
        ForestSites1,
        ForestSites2,
        ForestSubnets,
        ForestSubnets1,
        ForestSubnets2,
        ForestSiteLinks,

        DomainRootDSE,
        DomainAuthenticationPolicies, // Not yet tested
        DomainAuthenticationPolicySilos, // Not yet tested
        DomainCentralAccessPolicies, // Not yet tested
        DomainCentralAccessRules, // Not yet tested
        DomainClaimTransformPolicies, // Not yet tested
        DomainClaimTypes, // Not yet tested

        DomainFineGrainedPolicies,
        DomainGUIDS,
        DomainDNSSRV,
        DomainDNSA,

        DomainAdministrators,
        DomainAdministratorsRecursive,
        EnterpriseAdministrators,
        EnterpriseAdministratorsRecursive,
        DomainInformation,
        DomainControllers,
        DomainFSMO,
        DomainDefaultPasswordPolicy,
        DomainGroupPolicies,
        DomainGroupPoliciesDetails,
        DomainGroupPoliciesACL,
        DomainOrganizationalUnits,
        DomainOrganizationalUnitsBasicACL,
        DomainOrganizationalUnitsExtended,
        DomainContainers,
        DomainPriviligedGroupMembers,
        DomainUsersFullList,

        DomainUsers,
        DomainUsersAll,
        DomainUsersSystemAccounts,
        DomainUsersNeverExpiring,
        DomainUsersNeverExpiringInclDisabled,
        DomainUsersExpiredInclDisabled,
        DomainUsersExpiredExclDisabled,

        DomainUsersCount,
        DomainTrusts,

        DomainGroupsFullList,
        DomainGroups,
        DomainGroupsMembers,
        DomainGroupsMembersRecursive,

        DomainGroupsRest,
        DomainGroupsSpecial,
        DomainGroupsPriviliged,
        DomainGroupMembersRecursiveRest,
        DomainGroupMembersRecursiveSpecial,
        DomainGroupMembersRecursivePriviliged,

        DomainComputersFullList
    }
"@