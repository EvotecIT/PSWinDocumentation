# PSWinDocumentation - PowerShell Module

Overview of this module is available at: https://evotec.xyz/hub/scripts/pswindocumentation-powershell-module/

## This module utilizes three projects of mine:
- [PSWriteWord](https://evotec.xyz/hub/scripts/pswriteword-powershell-module/) - creating **Microsoft Wor**d without Word installed from PowerShell
- [PSWriteExcel](https://evotec.xyz/hub/scripts/pswriteexcel-powershell-module/) - creating **Microsoft Excel** without Excel installed from PowerShell
- [PSSharedGoods](https://github.com/EvotecIT/PSSharedGoods) - bunch of useful functions that I share among all of my projects

All 3 modules above are under active development.

## For additional functionality it utilizes:
- [AWSPowershell](https://www.powershellgallery.com/packages/AWSPowerShell/) - allows connecting to AWS and creating AWS dataset
- [ActiveDirectory RSAT](https://www.microsoft.com/en-us/download/details.aspx?id=45520) - allows connecting to AD and creating Active Directory dataset
- [DBATools](https://www.powershellgallery.com/packages/dbatools/) - allows connecting to SQL and sending data to SQL (create table, alter table, inserts)
- [DSInternals](https://www.powershellgallery.com/packages/DSInternals) - required for Password Audit in Active Directory

You have to install those modules **yourself**. I don't bundle them but it's as easy as typing `Install-Module <YourModule>`. While I could bundle them and maybe some day I will but for now it's too heavy as this module can be installed on AD servers directly - so I don't want you to overload your DC's - even if it's just PowerShell Module.

## Goals

Ultimate goal of this project is to have proper documentation of following services:

- Active Directory - **80% done** (more or less)
- AWS - **50% done** (more or less)
- Microsoft Exchange - **20% done** (mostly Excel / SQL export oriented)
- Office 365 Exchange Online - **20% done** (mostly Excel / SQL export oriented)
- Office 365 Azure AD - **20% don**e (mostly Excel / SQL export oriented)
- Office 365 Teams (not started)
- Office 365 Skype for Business (not started)
- Office 365 Intune (not started)
- Office 365 SharePoint (not started)
- Windows Server (some basic stuff - not functionable)
- Windows Workstation (some basic stuff - not functionable)

I'm heavily open for feature requests and people willing to create data sets. By data sets I mean data prepared to be included in report (usually HashTable, OrderedHashTable, Array or PSCustomobject). This module is smart enough that it can easily convert that data into Word Sections. Also don't really pay attention to percentage numbers. If you have request I will consier adding it.

## Updates
- 0.3.x / 2018.10.08 - [full blog post](https://evotec.xyz/pswindocumentation-audit-active-directory-passwords/)
    - Added **audit your Active Directory Passwords**
    - Expanded Active Directory data types (computer based)
    - Expanded Active Directory data types (fine grained policies)
    - Couple of fixes here and there
- 0.2.x / 2018.09.23 - [full blog post](https://evotec.xyz/pswindocumentation-export-to-word-excel-sql-of-ad-aws-exchange-o365-exchange-o365-azure-ad/)
    - Allows Exporting to Microsoft SQL (that's right – export data directly to SQL – complete with create table, alter table and inserts)
    - Basic data set AWS
    - Advanced data set Active Directory
    - Basic data set Microsoft Exchange
    - Basic data set Office 365 – Exchange Online
    - Basic data set Office 365 – Azure AD
    - Prescanning of data headers for exports (unravel hidden data)
    - Ability to define TableColumnWidths in sections

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

Following is **very incomplete list** of things that are done or are planned in near future. I really need to update that.

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
    DomainUsersCount,
    DomainUsersAll,
    DomainUsersSystemAccounts,
    DomainUsersNeverExpiring,
    DomainUsersNeverExpiringInclDisabled,
    DomainUsersExpiredInclDisabled,
    DomainUsersExpiredExclDisabled,
    DomainAdministrators,
    DomainAdministratorsRecursive,
    DomainEnterpriseAdministrators,
    DomainEnterpriseAdministratorsRecursive,

    // Domain Information - Computer Data
    DomainComputersFullList, // Contains all data
    DomainComputersAll,
    DomainComputersAllCount,
    DomainComputers,
    DomainComputersCount,
    DomainServers,
    DomainServersCount,
    DomainComputersUnknown,
    DomainComputersUnknownCount,

    // This requires DSInstall PowerShell Module
    DomainPasswordDataUsers, // Gathers users data and their passwords
    DomainPasswordDataPasswords, // Compares Users Password with File
    DomainPasswordDataPasswordsHashes, // Compares Users Password with File HASH
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
    DomainPasswordHashesWeakPassword,
    DomainPasswordStats,
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



### Statistics

## Known Issues

- [ ] If you want to build documentation on your own template you're free to do so, however you should use **Data\EmptyDocument.docx** as your starting template. This is because of issue with **PSWriteWord** (more specifically with DLL it uses) where Heading styles are not available when using template created directly in Microsoft Word. Issue was reported [PSWriteWord #16](https://github.com/EvotecIT/PSWriteWord/issues/16) but it's a long time till it will be fixed. Until then simply use **Data\EmptyDocument.docx** and then add your logos, text, whatever you feel like you need in template.