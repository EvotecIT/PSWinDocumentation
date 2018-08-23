Add-Type -TypeDefinition @"
    public enum ActiveDirectory {
        // Forest Information - Section Main
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

        // Domain Information - Section Main

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
        DomainTrusts,

        // Domain Information - Group Data
        DomainGroupsFullList, // Contains all data

        DomainGroups,
        DomainGroupsMembers,
        DomainGroupsMembersRecursive,

        DomainGroupsSpecial,
        DomainGroupsSpecialMembers,
        DomainGroupsSpecialMembersRecursive,

        DomainGroupsPriviliged,
        DomainGroupsPriviligedMembers,
        DomainGroupsPriviligedMembersRecursive,

        // Domain Information - User Data
        DomainUsersFullList, // Contains all data

        DomainUsers,
        DomainUsersAll,
        DomainUsersSystemAccounts,
        DomainUsersNeverExpiring,
        DomainUsersNeverExpiringInclDisabled,
        DomainUsersExpiredInclDisabled,
        DomainUsersExpiredExclDisabled,

        DomainUsersCount,

        DomainAdministrators,
        DomainAdministratorsRecursive,
        EnterpriseAdministrators,
        EnterpriseAdministratorsRecursive,


        // Domain Information - Computer Data
        DomainComputersFullList // Contains all data

    }
"@