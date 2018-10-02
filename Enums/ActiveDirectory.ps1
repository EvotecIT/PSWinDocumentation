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
        DomainRIDs,
        DomainAuthenticationPolicies, // Not yet tested
        DomainAuthenticationPolicySilos, // Not yet tested
        DomainCentralAccessPolicies, // Not yet tested
        DomainCentralAccessRules, // Not yet tested
        DomainClaimTransformPolicies, // Not yet tested
        DomainClaimTypes, // Not yet tested
        DomainFineGrainedPolicies,
        DomainFineGrainedPoliciesUsers,
        DomainFineGrainedPoliciesUsersExtended,
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

        DomainEnterpriseAdministrators,
        DomainEnterpriseAdministratorsRecursive,

        // Domain Information - Computer Data
        DomainComputersFullList, // Contains all data

        // This requires DSInstall PowerShell Module
        DomainPasswordClearTextPassword,
        DomainPasswordLMHash,
        DomainPasswordEmptyPassword,
        DomainPasswordWeakPassword,
        DomainPasswordDefaultComputerPassword,
        DomainPasswordPasswordNotRequired,
        DomainPasswordPasswordNeverExpires,
        DomainPasswordAESKeysMissing,
        DomainPasswordPreAuthNotRequired,
        DomainPasswordDESEncryptionOnly,
        DomainPasswordDelegatableAdmins,
        DomainPasswordDuplicatePasswordGroups,

        DomainPasswordHashesClearTextPassword,
        DomainPasswordHashesLMHash,
        DomainPasswordHashesEmptyPassword,
        DomainPasswordHashesWeakPassword,
        DomainPasswordHashesDefaultComputerPassword,
        DomainPasswordHashesPasswordNotRequired,
        DomainPasswordHashesPasswordNeverExpires,
        DomainPasswordHashesAESKeysMissing,
        DomainPasswordHashesPreAuthNotRequired,
        DomainPasswordHashesDESEncryptionOnly,
        DomainPasswordHashesDelegatableAdmins,
        DomainPasswordHashesDuplicatePasswordGroups,

        DomainPasswordStats,
        DomainPasswordHashesStats

    }
"@