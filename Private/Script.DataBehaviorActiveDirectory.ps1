$Script:DataBehaviorActiveDirectory = [ordered] @{
    ForestInformation                        = @{
        OnlineRequired = $true
    }
    ForestFSMO                               = @{
        OnlineRequired = $true
    }
    ForestGlobalCatalogs                     = @{
        OnlineRequired = $true
    }
    ForestOptionalFeatures                   = @{
        OnlineRequired = $true
    }
    ForestUPNSuffixes                        = @{
        OnlineRequired = $true
    }
    ForestSPNSuffixes                        = @{
        OnlineRequired = $true
    }
    ForestSites                              = @{
        OnlineRequired = $true
    }
    ForestSites1                             = @{
        OnlineRequired = $false
    }
    ForestSites2                             = @{
        OnlineRequired = $false
    }
    ForestSubnets                            = @{
        OnlineRequired = $true
    }
    ForestSubnets1                           = @{
        OnlineRequired = $true
    }
    ForestSubnets2                           = @{
        OnlineRequired = $true
    }
    ForestSiteLinks                          = @{
        OnlineRequired = $true
    }
    DomainRootDSE                            = @{
        OnlineRequired = $true
    }
    DomainRIDs                               = @{
        OnlineRequired = $true
    }
    DomainAuthenticationPolicies             = @{
        OnlineRequired = $true
    }
    DomainAuthenticationPolicySilos          = @{
        OnlineRequired = $true
    }
    DomainCentralAccessPolicies              = @{
        OnlineRequired = $true
    }
    DomainCentralAccessRules                 = @{
        OnlineRequired = $true
    }
    DomainClaimTransformPolicies             = @{
        OnlineRequired = $true
    }
    DomainClaimTypes                         = @{
        OnlineRequired = $true
    }
    DomainFineGrainedPolicies                = @{
        OnlineRequired = $true
    }
    DomainFineGrainedPoliciesUsers           = @{
        OnlineRequired = $true
    }
    DomainFineGrainedPoliciesUsersExtended   = @{
        OnlineRequired = $true
    }
    DomainGUIDS                              = @{
        OnlineRequired = $true
    }
    DomainDNSSRV                             = @{
        OnlineRequired = $true
    }
    DomainDNSA                               = @{
        OnlineRequired = $true
    }
    DomainInformation                        = @{
        OnlineRequired = $true
    }
    DomainControllers                        = @{
        OnlineRequired = $true
    }
    DomainFSMO                               = @{
        OnlineRequired = $true
    }
    DomainDefaultPasswordPolicy              = @{
        OnlineRequired = $true
    }
    DomainGroupPolicies                      = @{
        OnlineRequired = $true
    }
    DomainGroupPoliciesDetails               = @{
        OnlineRequired = $true
    }
    DomainGroupPoliciesACL                   = @{
        OnlineRequired = $true
    }
    DomainOrganizationalUnits                = @{
        OnlineRequired = $true
    }
    DomainOrganizationalUnitsBasicACL        = @{
        OnlineRequired = $true
    }
    DomainOrganizationalUnitsExtended        = @{
        OnlineRequired = $true
    }
    DomainContainers                         = @{
        OnlineRequired = $true
    }
    DomainTrusts                             = @{
        OnlineRequired = $true
    }

    DomainGroupsFullList                     = @{
        OnlineRequired = $true
    }

    DomainGroups                             = @{
        OnlineRequired = $true
    }
    DomainGroupsMembers                      = @{
        OnlineRequired = $true
    }
    DomainGroupsMembersRecursive             = @{
        OnlineRequired = $true
    }
    DomainGroupsSpecial                      = @{
        OnlineRequired = $true
    }
    DomainGroupsSpecialMembers               = @{
        OnlineRequired = $true
    }
    DomainGroupsSpecialMembersRecursive      = @{
        OnlineRequired = $true
    }

    DomainGroupsPriviliged                   = @{
        OnlineRequired = $true
    }
    DomainGroupsPriviligedMembers            = @{
        OnlineRequired = $true
    }
    DomainGroupsPriviligedMembersRecursive   = @{
        OnlineRequired = $true
    }

    DomainUsersFullList                      = @{
        OnlineRequired = $true
    }
    DomainUsers                              = @{
        OnlineRequired = $true
    }
    DomainUsersCount                         = @{
        OnlineRequired = $true
    }
    DomainUsersAll                           = @{
        OnlineRequired = $true
    }
    DomainUsersSystemAccounts                = @{
        OnlineRequired = $true
    }
    DomainUsersNeverExpiring                 = @{
        OnlineRequired = $true
    }
    DomainUsersNeverExpiringInclDisabled     = @{
        OnlineRequired = $true
    }
    DomainUsersExpiredInclDisabled           = @{
        OnlineRequired = $true
    }
    DomainUsersExpiredExclDisabled           = @{
        OnlineRequired = $true
    }
    DomainAdministrators                     = @{
        OnlineRequired = $true
    }
    DomainAdministratorsRecursive            = @{
        OnlineRequired = $true
    }
    DomainEnterpriseAdministrators           = @{
        OnlineRequired = $true
    }
    DomainEnterpriseAdministratorsRecursive  = @{
        OnlineRequired = $true
    }
    DomainComputersFullList                  = @{
        OnlineRequired = $true
    }
    DomainComputersAll                       = @{
        OnlineRequired = $true
    }
    DomainComputersAllCount                  = @{
        OnlineRequired = $true
    }
    DomainComputers                          = @{
        OnlineRequired = $true
    }
    DomainComputersCount                     = @{
        OnlineRequired = $true
    }
    DomainServers                            = @{
        OnlineRequired = $true
    }
    DomainServersCount                       = @{
        OnlineRequired = $true
    }
    DomainComputersUnknown                   = @{
        OnlineRequired = $true
    }
    DomainComputersUnknownCount              = @{
        OnlineRequired = $true
    }

    DomainPasswordDataUsers                  = @{
        OnlineRequired = $true
    }
    DomainPasswordDataPasswords              = @{
        OnlineRequired = $true
    }
    DomainPasswordDataPasswordsHashes        = @{
        OnlineRequired = $true
    }
    DomainPasswordClearTextPassword          = @{
        OnlineRequired = $true
    }
    DomainPasswordClearTextPasswordEnabled   = @{
        OnlineRequired = $true
    }
    DomainPasswordClearTextPasswordDisabled  = @{
        OnlineRequired = $true
    }
    DomainPasswordLMHash                     = @{
        OnlineRequired = $true
    }
    DomainPasswordEmptyPassword              = @{
        OnlineRequired = $true
    }
    DomainPasswordWeakPassword               = @{
        OnlineRequired = $true
    }
    DomainPasswordWeakPasswordEnabled        = @{
        OnlineRequired = $true
    }
    DomainPasswordWeakPasswordDisabled       = @{
        OnlineRequired = $true
    }
    DomainPasswordWeakPasswordList           = @{
        OnlineRequired = $true
    }
    DomainPasswordDefaultComputerPassword    = @{
        OnlineRequired = $true
    }
    DomainPasswordPasswordNotRequired        = @{
        OnlineRequired = $true
    }
    DomainPasswordPasswordNeverExpires       = @{
        OnlineRequired = $true
    }
    DomainPasswordAESKeysMissing             = @{
        OnlineRequired = $true
    }
    DomainPasswordPreAuthNotRequired         = @{
        OnlineRequired = $true
    }
    DomainPasswordDESEncryptionOnly          = @{
        OnlineRequired = $true
    }
    DomainPasswordDelegatableAdmins          = @{
        OnlineRequired = $true
    }
    DomainPasswordDuplicatePasswordGroups    = @{
        OnlineRequired = $true
    }
    DomainPasswordHashesWeakPassword         = @{
        OnlineRequired = $true
    }
    DomainPasswordHashesWeakPasswordEnabled  = @{
        OnlineRequired = $true
    }
    DomainPasswordHashesWeakPasswordDisabled = @{
        OnlineRequired = $true
    }
    DomainPasswordStats                      = @{
        OnlineRequired = $true
    }
}
