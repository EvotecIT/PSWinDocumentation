Import-Module PSWriteWord # -Force
#Import-Module PSWriteExcel
Import-Module PSWinDocumentation -Force
Import-Module PSWriteColor
Import-Module ActiveDirectory

$Document = [ordered]@{
    Configuration    = [ordered] @{
        Prettify       = @{
            CompanyName        = 'Evotec'
            UseBuiltinTemplate = $true
            CustomTemplatePath = ''
            Language           = 'en-US'
        }
        Options        = @{
            OpenDocument = $false
            OpenExcel    = $false
        }
        DisplayConsole = @{
            ShowTime   = $false
            LogFile    = "$ENV:TEMP\PSWinDocumentationTesting.log"
            TimeFormat = 'yyyy-MM-dd HH:mm:ss'
        }
        Debug          = @{
            Verbose = $false
        }
    }
    DocumentAD       = [ordered] @{
        Enable        = $true
        ExportWord    = $false
        ExportExcel   = $false
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWinDocumentation-Report.docx"
        FilePathExcel = "$Env:USERPROFILE\Desktop\PSWinDocumentation-Report.xlsx"
        Sections      = [ordered] @{
            SectionForest = [ordered] @{
                <#
                SectionTOC                = [ordered] @{
                    Use                  = $true
                    TocGlobalDefinition  = $true
                    TocGlobalTitle       = 'Table of content'
                    TocGlobalRightTabPos = 15
                    #TocGlobalSwitches    = 'A', 'C' #[TableContentSwitches]::C, [TableContentSwitches]::A
                    PageBreaksAfter      = 1
                }
                SectionForestIntroduction = [ordered] @{
                    ### Enables section
                    Use             = $true

                    ### Decides how TOC should be visible
                    TocEnable       = $True
                    TocText         = 'Scope'
                    TocListLevel    = 0
                    TocListItemType = [ListItemType]::Numbered
                    TocHeadingType  = [HeadingType]::Heading1

                    ### Text is added before table/list
                    Text            = "This document provides a low-level design of roles and permissions for" `
                        + " the IT infrastructure team at <CompanyName> organization. This document utilizes knowledge from" `
                        + " AD General Concept document that should be delivered with this document. Having all the information" `
                        + " described in attached document one can start designing Active Directory with those principles in mind." `
                        + " It's important to know while best practices that were described are important in decision making they" `
                        + " should not be treated as final and only solution. Most important aspect is to make sure company has full" `
                        + " usability of Active Directory and is happy with how it works. Making things harder just for the sake of" `
                        + " implementation of best practices isn't always the best way to go."
                    TextAlignment   = [Alignment]::Both
                    PageBreaksAfter = 1

                }
                #>
                SectionForestSummary = [ordered] @{
                    Use               = $true
                    TocEnable         = $True
                    TocText           = 'General Information - Forest Summary'
                    TocListLevel      = 0
                    TocListItemType   = [ListItemType]::Numbered
                    TocHeadingType    = [HeadingType]::Heading1
                    TableData         = [ActiveDirectory]::ForestInformation
                    TableDesign       = [TableDesign]::ColorfulGridAccent5
                    TableTitleMerge   = $true
                    TableTitleText    = "Forest Summary"
                    Text              = "Active Directory at <CompanyName> has a forest name <ForestName>." `
                        + " Following table contains forest summary with important information:"
                    ExcelExport       = $true
                    ExcelWorkSheet    = 'Forest Summary'
                    ExcelData         = [ActiveDirectory]::ForestInformation

                    SqlExport         = $true
                    SqlServer         = 'EVO1'
                    SqlDatabase       = 'SSAE18'
                    SqlData           = [ActiveDirectory]::ForestInformation
                    SqlTable          = 'dbo.[ForestInformation]'
                    SqlTableTranspose = $true
                    SqlTableCreate    = $true
                }


                SectionForestFSMO    = [ordered] @{
                    Use                   = $true
                    TableData             = [ActiveDirectory]::ForestFSMO
                    TableDesign           = 'ColorfulGridAccent5'
                    TableTitleMerge       = $true
                    TableTitleText        = 'FSMO Roles'
                    Text                  = 'Following table contains FSMO servers'
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $true
                    ExcelWorkSheet        = 'Forest FSMO'
                    ExcelData             = [ActiveDirectory]::ForestFSMO


                    SqlExport             = $true
                    SqlServer             = 'EVO1'
                    SqlDatabase           = 'SSAE18'
                    SqlData               = [ActiveDirectory]::ForestFSMO
                    SqlTable              = 'dbo.[ForestFSMO]'
                    SqlTableTranspose     = $true
                    SqlTableCreate        = $true
                }
                <#
                SectionForestOptionalFeatures = [ordered] @{
                    Use                   = $true
                    TableData             = [ActiveDirectory]::ForestOptionalFeatures
                    TableDesign           = [TableDesign]::ColorfulGridAccent5
                    TableTitleMerge       = $true
                    TableTitleText        = 'Optional Features'
                    Text                  = 'Following table contains optional forest features'
                    TextNoData            = "Following section should have table containing forest features. However no data was provided."
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $true
                    ExcelWorkSheet        = 'Forest Optional Features'
                    ExcelData             = [ActiveDirectory]::ForestOptionalFeatures
                }
                SectionForestUPNSuffixes      = [ordered] @{
                    Use                   = $true
                    Text                  = "Following UPN suffixes were created in this forest:"
                    TextNoData            = "No UPN suffixes were created in this forest."
                    ListType              = 'Bulleted'
                    ListData              = [ActiveDirectory]::ForestUPNSuffixes
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $true
                    ExcelWorkSheet        = 'Forest UPN Suffixes'
                    ExcelData             = [ActiveDirectory]::ForestUPNSuffixes
                }
                SectionForesSPNSuffixes       = [ordered] @{
                    Use                   = $true
                    Text                  = "Following SPN suffixes were created in this forest:"
                    TextNoData            = "No SPN suffixes were created in this forest."
                    ListType              = 'Bulleted'
                    ListData              = [ActiveDirectory]::ForestSPNSuffixes
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $true
                    ExcelWorkSheet        = 'Forest SPN Suffixes'
                    ExcelData             = [ActiveDirectory]::ForestSPNSuffixes
                }
                SectionForestSites1           = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Sites'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading1'
                    TableData       = [ActiveDirectory]::ForestSites1
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = "Forest Sites list can be found below"
                    ExcelExport     = $false  ## Exported as one below
                    ExcelWorkSheet  = 'Forest Sites 1'
                    ExcelData       = [ActiveDirectory]::ForestSites1
                }
                SectionForestSites2           = [ordered] @{
                    Use                   = $true
                    TableData             = [ActiveDirectory]::ForestSites2
                    TableDesign           = 'ColorfulGridAccent5'
                    Text                  = "Forest Sites list can be found below"
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $false ## Exported as one below
                    ExcelWorkSheet        = 'Forest Sites 2'
                    ExcelData             = [ActiveDirectory]::ForestSites2
                }
                SectionForestSites            = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = 'Forest Sites'
                    ExcelData      = [ActiveDirectory]::ForestSites
                }
                SectionForestSubnets1         = [ordered] @{
                    Use                   = $true
                    TocEnable             = $True
                    TocText               = 'General Information - Subnets'
                    TocListLevel          = 1
                    TocListItemType       = 'Numbered'
                    TocHeadingType        = 'Heading1'
                    TableData             = [ActiveDirectory]::ForestSubnets1
                    TableDesign           = 'ColorfulGridAccent5'
                    Text                  = "Table below contains information regarding relation between Subnets and sites"
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $true
                    ExcelWorkSheet        = 'Forest Subnets 1'
                    ExcelData             = [ActiveDirectory]::ForestSubnets1
                }
                SectionForestSubnets2         = [ordered] @{
                    Use                   = $true
                    TableData             = [ActiveDirectory]::ForestSubnets2
                    TableDesign           = 'ColorfulGridAccent5'
                    Text                  = "Table below contains information regarding relation between Subnets and sites"
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $true
                    ExcelWorkSheet        = 'Forest Subnets 2'
                    ExcelData             = [ActiveDirectory]::ForestSubnets2
                }
                SectionForestSiteLinks        = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Site Links'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading1'
                    TableData       = [ActiveDirectory]::ForestSiteLinks
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = "Forest Site Links information is available in table below"
                    ExcelExport     = $true
                    ExcelWorkSheet  = 'Forest Site Links'
                    ExcelData       = [ActiveDirectory]::ForestSiteLinks
                }
                #>
            }
            SectionDomain = [ordered] @{
                <#
                SectionPageBreak                                  = [ordered] @{
                    Use              = $True
                    PageBreaksBefore = 1
                }
                SectionDomainStarter                              = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Domain <Domain>'
                    TocListLevel    = 0
                    TocListItemType = [ListItemType]::Numbered
                    TocHeadingType  = [HeadingType]::Heading1
                }
                SectionDomainIntroduction                         = [ordered] @{
                    Use                   = $true
                    TocEnable             = $True
                    TocText               = 'General Information - Domain Summary'
                    TocListLevel          = 1
                    TocListItemType       = [ListItemType]::Numbered
                    TocHeadingType        = [HeadingType]::Heading1
                    Text                  = "Following domain exists within forest <ForestName>:"
                    ListBuilderContent    = "Domain <DomainDN>", 'Name for fully qualified domain name (FQDN): <Domain>', 'Name for NetBIOS: <DomainNetBios>'
                    ListBuilderLevel      = 0, 1, 1
                    ListBuilderType       = [ListItemType]::Bulleted, [ListItemType]::Bulleted, [ListItemType]::Bulleted
                    EmptyParagraphsBefore = 0
                }
                SectionDomainControllers                          = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'General Information - Domain Controllers'
                    TocListLevel        = 1
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainControllers
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 8
                    Text                = 'Following table contains domain controllers'
                    TextNoData          = ''
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - DCs'
                    ExcelData           = [ActiveDirectory]::DomainControllers
                }
                SectionDomainFSMO                                 = [ordered] @{
                    Use                   = $true
                    TableData             = [ActiveDirectory]::DomainFSMO
                    TableDesign           = 'ColorfulGridAccent5'
                    TableTitleMerge       = $true
                    TableTitleText        = "FSMO Roles for <Domain>"
                    Text                  = "Following table contains FSMO servers with roles for domain <Domain>"
                    EmptyParagraphsBefore = 1
                    ExcelExport           = $true
                    ExcelWorkSheet        = '<Domain> - FSMO'
                    ExcelData             = [ActiveDirectory]::DomainFSMO
                }
                SectionDomainDefaultPasswordPolicy                = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Password Policies'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainDefaultPasswordPolicy
                    TableDesign     = 'ColorfulGridAccent5'
                    TableTitleMerge = $True
                    TableTitleText  = "Default Password Policy for <Domain>"
                    Text            = 'Following table contains password policies for all users within <Domain>'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - DefaultPasswordPolicy'
                    ExcelData       = [ActiveDirectory]::DomainDefaultPasswordPolicy
                }
                SectionDomainFineGrainedPolicies                  = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'General Information - Fine Grained Password Policies'
                    TocListLevel        = 1
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainFineGrainedPolicies
                    TableDesign         = [TableDesign]::ColorfulGridAccent5
                    TableMaximumColumns = 8
                    TableTitleMerge     = $false
                    TableTitleText      = "Fine Grained Password Policy for <Domain>"
                    Text                = 'Following table contains fine grained password policies'
                    TextNoData          = "Following section should cover fine grained password policies. " `
                        + "There were no fine grained password polices defined in <Domain>. There was no formal requirement to have " `
                        + "them set up."
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - Password Policy (Grained)'
                    ExcelData           = [ActiveDirectory]::DomainFineGrainedPolicies
                }
                SectionDomainGroupPolicies                        = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Group Policies'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainGroupPolicies
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = "Following table contains group policies for <Domain>"
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - GroupPolicies'
                    ExcelData       = [ActiveDirectory]::DomainGroupPolicies
                }
                SectionDomainGroupPoliciesDetails                 = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'General Information - Group Policies Details'
                    TocListLevel        = 1
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainGroupPoliciesDetails
                    TableMaximumColumns = 6
                    TableDesign         = 'ColorfulGridAccent5'
                    Text                = "Following table contains group policies for <Domain>"
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - GroupPolicies Details'
                    ExcelData           = [ActiveDirectory]::DomainGroupPoliciesDetails
                }
                SectionDomainGroupPoliciesACL                     = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'General Information - Group Policies ACL'
                    TocListLevel        = 1
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainGroupPoliciesACL
                    TableMaximumColumns = 6
                    TableDesign         = 'ColorfulGridAccent5'
                    Text                = "Following table contains group policies ACL for <Domain>"
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - GroupPoliciesACL'
                    ExcelData           = [ActiveDirectory]::DomainGroupPoliciesACL
                }
                SectionDomainDNSSrv                               = [ordered] @{
                    Use                  = $true
                    TocEnable            = $True
                    TocText              = 'General Information - DNS A/SRV Records'
                    TocListLevel         = 1
                    TocListItemType      = 'Numbered'
                    TocHeadingType       = 'Heading2'
                    TableData            = [ActiveDirectory]::DomainDNSSRV
                    TableMaximumColumns  = 10
                    TableDesign          = 'ColorfulGridAccent5'
                    Text                 = "Following table contains SRV records for Kerberos and LDAP"
                    EmptyParagraphsAfter = 1
                    ExcelExport          = $true
                    ExcelWorkSheet       = '<Domain> - DNSSRV'
                    ExcelData            = [ActiveDirectory]::DomainDNSSRV
                }
                SectionDomainDNSA                                 = [ordered] @{
                    Use                 = $true
                    TableData           = [ActiveDirectory]::DomainDNSA
                    TableMaximumColumns = 10
                    TableDesign         = 'ColorfulGridAccent5'
                    Text                = "Following table contains A records for Kerberos and LDAP"
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - DNSA'
                    ExcelData           = [ActiveDirectory]::DomainDNSA
                }
                SectionDomainTrusts                               = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'General Information - Trusts'
                    TocListLevel        = 1
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainTrusts
                    TableMaximumColumns = 6
                    TableDesign         = 'ColorfulGridAccent5'
                    Text                = "Following table contains trusts established with domains..."
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - DomainTrusts'
                    ExcelData           = [ActiveDirectory]::DomainTrusts
                }
                SectionDomainOrganizationalUnits                  = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Organizational Units'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainOrganizationalUnits
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = "Following table contains all OU's created in <Domain>"
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - OrganizationalUnits'
                    ExcelData       = [ActiveDirectory]::DomainOrganizationalUnits
                }
                SectionDomainPriviligedGroup                      = [ordered] @{
                    Use             = $False
                    TocEnable       = $True
                    TocText         = 'General Information - Priviliged Groups'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainGroupsPriviliged
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = 'Following table contains list of priviliged groups and count of the members in it.'
                    ChartEnable     = $True
                    ChartTitle      = 'Priviliged Group Members'
                    ChartData       = [ActiveDirectory]::DomainGroupsPriviliged
                    ChartKeys       = 'Group Name', 'Members Count'
                    ChartValues     = 'Members Count'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - PriviligedGroupMembers'
                    ExcelData       = [ActiveDirectory]::DomainGroupsPriviliged
                }
                SectionDomainAdministrators                       = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Domain Administrators'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainAdministratorsRecursive
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = 'Following users have highest priviliges and are able to control a lot of Windows resources.'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - DomainAdministrators'
                    ExcelData       = [ActiveDirectory]::DomainAdministratorsRecursive
                }
                SectionEnterpriseAdministrators                   = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Enterprise Administrators'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainEnterpriseAdministratorsRecursive
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = 'Following users have highest priviliges across Forest and are able to control a lot of Windows resources.'
                    TextNoData      = 'No Enterprise Administrators users were defined for this domain.'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - EnterpriseAdministrators'
                    ExcelData       = [ActiveDirectory]::DomainEnterpriseAdministratorsRecursive
                }
                SectionDomainUsersCount                           = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Users Count'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainUsersCount
                    TableDesign     = 'ColorfulGridAccent5'
                    TableTitleMerge = $False
                    TableTitleText  = 'Users Count'
                    Text            = "Following table and chart shows number of users in its categories"
                    ChartEnable     = $True
                    ChartTitle      = 'Users Count'
                    ChartData       = [ActiveDirectory]::DomainUsersCount
                    ChartKeys       = 'Keys'
                    ChartValues     = 'Values'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - UsersCount'
                    ExcelData       = [ActiveDirectory]::DomainUsersCount
                }
                SectionExcelDomainOrganizationalUnitsBasicACL     = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - OU ACL Basic'
                    ExcelData      = [ActiveDirectory]::DomainOrganizationalUnitsBasicACL
                }
                SectionExcelDomainOrganizationalUnitsExtended     = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - OU ACL Extended'
                    ExcelData      = [ActiveDirectory]::DomainOrganizationalUnitsExtended
                }
                #>
                SectionExcelDomainUsers               = [ordered] @{
                    Use              = $true
                    ExcelExport      = $false
                    ExcelWorkSheet   = '<Domain> - Users'
                    ExcelData        = [ActiveDirectory]::DomainUsers

                    SqlExport        = $true
                    SqlServer        = 'EVO1'
                    SqlDatabase      = 'SSAE18'
                    SqlData          = [ActiveDirectory]::DomainUsers
                    SqlTable         = 'dbo.[Users]'
                    SqlTableCreate   = $true
                    SqlTabl1eMapping = [ordered] @{
                        # Left Side is data in PSWinReporting
                        # Right Side is column name in SQL
                        # Changing makes sense only for left side...
                        # Use this if you need to have different mapping
                        'Name'                              = 'Name'
                        'UserPrincipalName'                 = 'UserPrincipalName'
                        'SamAccountName'                    = 'SamAccountName'
                        'Display Name'                      = 'Display Name'
                        'Given Name'                        = 'Given Name'
                        'Surname'                           = 'Surname'
                        'EmailAddress'                      = 'EmailAddress'
                        'PasswordExpired'                   = 'PasswordExpired'
                        'PasswordLastSet'                   = 'PasswordLastSet'
                        'PasswordNotRequired'               = 'PasswordNotRequired'
                        'PasswordNeverExpires'              = 'PasswordNeverExpires'
                        'Enabled'                           = 'Enabled'
                        'Manager'                           = 'Manager'
                        'Manager Email'                     = 'Manager Email'
                        'DateExpiry'                        = 'DateExpiry'
                        'DaysToExpire'                      = 'DaysToExpire'
                        'AccountExpirationDate'             = 'AccountExpirationDate'
                        'AccountLockoutTime'                = 'AccountLockoutTime'
                        'AllowReversiblePasswordEncryption' = 'AllowReversiblePasswordEncryption'
                        'BadLogonCount'                     = 'BadLogonCount'
                        'CannotChangePassword'              = 'CannotChangePassword'
                        'CanonicalName'                     = 'CanonicalName'
                        'Description'                       = 'Description'
                        'DistinguishedName'                 = 'DistinguishedName'
                        'EmployeeID'                        = 'EmployeeID'
                        'EmployeeNumber'                    = 'EmployeeNumber'
                        'LastBadPasswordAttempt'            = 'LastBadPasswordAttempt'
                        'LastLogonDate'                     = 'LastLogonDate'
                        'Created'                           = 'Created'
                        'Modified'                          = 'Modified'
                        'Protected'                         = 'Protected'
                        'Primary Group'                     = 'Primary Group'
                        'Member Of'                         = 'Member Of'
                        'AddedWhen'                         = 'AddedWhen'# ColumnsToTrack when it was added to database and by who / not part of event
                        'AddedWho'                          = 'AddedWho'   # ColumnsToTrack when it was added to database and by who / not part of event
                    }
                }

                SectionExcelDomainUsersAll            = [ordered] @{
                    Use            = $false
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Users All'
                    ExcelData      = [ActiveDirectory]::DomainUsersAll
                }
                SectionExcelDomainUsersSystemAccounts = [ordered] @{
                    Use            = $false
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Users System'
                    ExcelData      = [ActiveDirectory]::DomainUsersSystemAccounts
                }
                SectionExcelDomainUsersNeverExpiring  = [ordered] @{
                    Use            = $true
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Never Expiring'
                    ExcelData      = [ActiveDirectory]::DomainUsersNeverExpiring

                    SqlExport      = $true
                    SqlServer      = 'EVO1'
                    SqlDatabase    = 'SSAE18'
                    SqlTableCreate = $true
                    SqlData        = [ActiveDirectory]::DomainUsersNeverExpiring
                    SqlTable       = 'dbo.[UsersNeverExpiring]'
                }
                <#
                SectionExcelDomainUsersNeverExpiringInclDisabled  = [ordered] @{
                    Use            = $true
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Never Expiring incl Disabled'
                    ExcelData      = [ActiveDirectory]::DomainUsersNeverExpiringInclDisabled
                }
                SectionExcelDomainUsersExpiredInclDisabled        = [ordered] @{
                    Use            = $true
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Expired incl Disabled'
                    ExcelData      = [ActiveDirectory]::DomainUsersExpiredInclDisabled
                }
                SectionExcelDomainUsersExpiredExclDisabled        = [ordered] @{
                    Use            = $true
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Expired excl Disabled'
                    ExcelData      = [ActiveDirectory]::DomainUsersExpiredExclDisabled
                }
                SectionExcelDomainUsersFullList                   = [ordered] @{
                    Use            = $true
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Users List Full'
                    ExcelData      = [ActiveDirectory]::DomainUsersFullList
                }
                SectionExcelDomainComputersFullList               = [ordered] @{
                    Use            = $true
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Computers List'
                    ExcelData      = [ActiveDirectory]::DomainComputersFullList
                }
                SectionExcelDomainGroupsFullList                  = [ordered] @{
                    Use            = $true
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Groups List'
                    ExcelData      = [ActiveDirectory]::DomainGroupsFullList
                }
                SectionExcelDomainGroupsRest                      = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Groups'
                    ExcelData      = [ActiveDirectory]::DomainGroups
                }
                #>

                SectionExcelDomainGroupsSpecial       = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Groups Special'
                    ExcelData      = [ActiveDirectory]::DomainGroupsSpecial

                    SqlExport      = $true
                    SqlServer      = 'EVO1'
                    SqlDatabase    = 'SSAE18'
                    SqlData        = [ActiveDirectory]::DomainGroupsSpecial
                    SqlTableCreate = $true
                    SqlTable       = 'dbo.[DomainGroupsSpecial]'
                }
                <#
                SectionExcelDomainGroupsPriviliged                = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Groups Priv'
                    ExcelData      = [ActiveDirectory]::DomainGroupsPriviliged
                }
                SectionExcelDomainGroupMembers                    = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Members'
                    ExcelData      = [ActiveDirectory]::DomainGroupsMembers
                }
                SectionExcelDomainGroupMembersSpecial             = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Members Special'
                    ExcelData      = [ActiveDirectory]::DomainGroupsSpecialMembers
                }
                SectionExcelDomainGroupMembersPriviliged          = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Members Priv'
                    ExcelData      = [ActiveDirectory]::DomainGroupsPriviligedMembers
                }
                SectionExcelDomainGroupMembersRecursive           = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Members Rec'
                    ExcelData      = [ActiveDirectory]::DomainGroupsMembersRecursive
                }
                SectionExcelDomainGroupMembersSpecialRecursive    = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Members RecSpecial'
                    ExcelData      = [ActiveDirectory]::DomainGroupsSpecialMembersRecursive
                }
                SectionExcelDomainGroupMembersPriviligedRecursive = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Members RecPriv'
                    ExcelData      = [ActiveDirectory]::DomainGroupsPriviligedMembersRecursive
                }
                #>
            }
        }

    }
    DocumentExchange = [ordered] @{

    }
}

Start-Documentation -Document $Document -Verbose