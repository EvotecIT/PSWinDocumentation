# PSWinDocumentation - PowerShell Module

Overview of this module is available at: https://evotec.xyz/hub/scripts/pswindocumentation-powershell-module/

This module utilizes two projects of mine:
- [PSWriteWord](https://evotec.xyz/hub/scripts/pswriteword-powershell-module/)
- [PSWriteExcel](https://evotec.xyz/hub/scripts/pswriteexcel-powershell-module/)

For additional functionality it utilizes:
- [AWSPowershell](https://www.powershellgallery.com/packages/AWSPowerShell/) - **work in progress** - not yet ready
- ActiveDirectory module - available as part of [RSAT](https://www.microsoft.com/en-us/download/details.aspx?id=45520)

 Both are under development and going step by step on per need basis.

## Goals

Ultimate goal of this project is to have proper documentation of following services:

- Active Directory
- Microsoft Exchange
- Office 365
- Windows Server
- Windows Workstation

I'm heavily open for feature requests and people willing to create data sets. By data sets I mean data prepared to be included in report (usually HashTable, OrderedHashTable, Array or PSCustomobject). This module is smart enough that it can easily convert that data into Word Sections.

## Updates
- 0.1 / 2018.08.23
    - Large release
    - You can read about it in separate [blog post](https://evotec.xyz/pswindocumentation-version-0-1-with-word-excel-export/)
    - Watch about it on [YouTube](https://youtu.be/6Vr3hEo2510) and [YouTube](https://youtu.be/c2kD_duHgTw)
- 0.0.5 / 2018.07.30
    -  fix for DefaultPassWordPolicy MinPasswordLength (was reporting wrong value)
- 0.0.4 / 2018.07.30
    -  added domain controllers section
    -  added few verbose messages with -Verbose switch for easier debugging
    -  commented out some unused code for now (to speed up work)
- 0.0.3 / 2018.07.29
    - first "good" release

## Progress on Documentation

### Active Directory

Following is incomplete list of things that are done or are planned in near future.

- [x] Forest Summary
- [x] Forest FSMO Roles
- [x] Forest Optional Features (Recycle Bin, PAM)
- [x] Forest UPN List
- [x] Forest SPN List
- [x] Domain Summary
- [ ] Domain Controllers
    - [X] Basic information
    - [ ] Basic hardware information
- [x] Domain FSMO Roles
- [x] Domain Password Policies
- [x] Domain Group Policies
- [ ] Domain Organizational Units
    - [ ] Requires work. Currently a bit useless
- [x] Domain Privileged Members (Groups)
- [x] Domain Administrators (All users)
- [x] Domain User Count
    - [X] Users Count Incl. System
    - [X] Users Count
    - [X] Users Expired
    - [X] Users Expired Incl. Disabled
    - [X] Users Never Expiring
    - [X] Users Never Expiring Incl. Disabled
    - [X] Users System Accounts
- [ ] Domain User List (deciding if needed)
    - [ ] Users Count Incl. System
    - [ ] Users Count
    - [ ] Users Expired
    - [ ] Users Expired Incl. Disabled
    - [ ] Users Never Expiring
    - [ ] Users Never Expiring Incl. Disabled
    - [ ] Users System Accounts

### Active Directory Data Sources - to use with new version

```powershell
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
DomainEnterpriseAdministrators,
DomainEnterpriseAdministratorsRecursive,
// Domain Information - Computer Data
DomainComputersFullList // Contains all data
 }
```



More suggestions are welcome.

### Microsoft Exchange
- [ ] Currently not defined. Feel free to make feature requests

### Microsoft Office 365
- [ ] Currently not defined. Feel free to make feature requests

### Windows Server doc
- [ ] Currently not defined. Feel free to make feature requests

### Windows Workstation doc
- [ ] Currently not defined. Feel free to make feature requests