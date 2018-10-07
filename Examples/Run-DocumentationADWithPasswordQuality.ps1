Import-Module PSWriteWord
Import-Module PSWriteExcel
Import-Module PSWinDocumentation #-Force
Import-Module PSWriteColor
Import-Module PSSharedGoods #-Force
Import-Module DbaTools
Import-Module ActiveDirectory
Import-Module AWSPowerShell

$Document = [ordered]@{
    Configuration = [ordered] @{
        Prettify       = @{
            CompanyName        = 'Evotec'
            UseBuiltinTemplate = $true
            CustomTemplatePath = ''
            Language           = 'en-US'
        }
        Options        = @{
            OpenDocument = $true
            OpenExcel    = $true
        }
        DisplayConsole = @{
            ShowTime   = $false
            LogFile    = "$ENV:TEMP\PSWinDocumentationADTesting.log"
            TimeFormat = 'yyyy-MM-dd HH:mm:ss'
        }
        Debug          = @{
            Verbose = $false
        }
    }
    DocumentAD    = [ordered] @{
        Enable        = $true
        ExportWord    = $true
        ExportExcel   = $true
        ExportSql     = $false
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ADReportWithPasswords.docx"
        FilePathExcel = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ADReportWithPasswords.xlsx"
        Configuration = [ordered] @{
            PasswordTests = @{
                Use                       = $true
                PasswordFilePathClearText = 'C:\Users\pklys\OneDrive - Evotec\Support\GitHub\PSWinDocumentation\Ignore\Passwords.txt'
                # Fair warning it will take ages if you use HaveIBeenPwned DB :-)
                UseHashDB                 = $true
                PasswordFilePathHash      = 'C:\Users\pklys\Downloads\pwned-passwords-ntlm-ordered-by-count\pwned-passwords-ntlm-ordered-by-count.txt'
            }
        }
        Sections      = [ordered] @{
            SectionForest = [ordered] @{
                SectionTOC                    = [ordered] @{
                    Use                  = $true
                    TocGlobalDefinition  = $true
                    TocGlobalTitle       = 'Table of content'
                    TocGlobalRightTabPos = 15
                    #TocGlobalSwitches    = 'A', 'C' #[TableContentSwitches]::C, [TableContentSwitches]::A
                    PageBreaksAfter      = 1
                }
                SectionForestIntroduction     = [ordered] @{
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
                SectionForestSummary          = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Forest Summary'
                    TocListLevel    = 0
                    TocListItemType = [ListItemType]::Numbered
                    TocHeadingType  = [HeadingType]::Heading1
                    TableData       = [ActiveDirectory]::ForestInformation
                    TableDesign     = [TableDesign]::ColorfulGridAccent5
                    TableTitleMerge = $true
                    TableTitleText  = "Forest Summary"
                    Text            = "Active Directory at <CompanyName> has a forest name <ForestName>." `
                        + " Following table contains forest summary with important information:"
                    ExcelExport     = $true
                    ExcelWorkSheet  = 'Forest Summary'
                    ExcelData       = [ActiveDirectory]::ForestInformation
                }
                SectionForestFSMO             = [ordered] @{
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
                }
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
            }
            SectionDomain = [ordered] @{
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
                    Use            = $true
                    #TocEnable           = $True
                    #TocText             = 'General Information - Group Policies ACL'
                    #TocListLevel        = 1
                    #TocListItemType     = 'Numbered'
                    #TocHeadingType      = 'Heading2'
                    #TableData           = [ActiveDirectory]::DomainGroupPoliciesACL
                    #TableMaximumColumns = 6
                    #TableDesign         = 'ColorfulGridAccent5'
                    #Text                = "Following table contains group policies ACL for <Domain>"
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - GroupPoliciesACL'
                    ExcelData      = [ActiveDirectory]::DomainGroupPoliciesACL
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
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'General Information - Organizational Units'
                    TocListLevel        = 1
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainOrganizationalUnits
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = "Following table contains all OU's created in <Domain>"
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - OrganizationalUnits'
                    ExcelData           = [ActiveDirectory]::DomainOrganizationalUnits
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
                SectionDomainUsers                                = [ordered] @{
                    Use              = $true
                    TocEnable        = $True
                    TocText          = 'General Information - Domain Users in <Domain>'
                    TocListLevel     = 1
                    TocListItemType  = [ListItemType]::Numbered
                    TocHeadingType   = [HeadingType]::Heading1
                    PageBreaksBefore = 1
                    Text             = 'Following section covers users information for domain <Domain>. '
                }
                SectionDomainUsersCount                           = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Users Count'
                    TocListLevel    = 2
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainUsersCount
                    TableDesign     = 'ColorfulGridAccent5'
                    TableTitleMerge = $true
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
                SectionDomainAdministrators                       = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Domain Administrators'
                    TocListLevel    = 2
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
                    TocListLevel    = 2
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
                SectionDomainComputers                            = [ordered] @{
                    Use              = $true
                    TocEnable        = $True
                    TocText          = 'General Information - Computer Objects in <Domain>'
                    TocListLevel     = 1
                    TocListItemType  = [ListItemType]::Numbered
                    TocHeadingType   = [HeadingType]::Heading1
                    PageBreaksBefore = 1
                    Text             = 'Following section covers computers information for domain <Domain>. '
                }
                DomainComputers                                   = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Computers'
                    TocListLevel    = 2
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainComputers
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = 'Following client computers are created in <Domain>.'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - DomainComputers'
                    ExcelData       = [ActiveDirectory]::DomainComputers
                }
                DomainComputersCount                              = [ordered] @{
                    Use             = $true
                    TableData       = [ActiveDirectory]::DomainComputersCount
                    TableDesign     = 'ColorfulGridAccent5'
                    TableTitleMerge = $true
                    TableTitleText  = 'Computers Count'
                    Text            = "Following table and chart shows number of computers and their versions"
                    ChartEnable     = $True
                    ChartTitle      = 'Computers Count'
                    ChartData       = [ActiveDirectory]::DomainComputersCount
                    ChartKeys       = 'System Name','System Count'
                    ChartValues     = 'System Count'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - DomainComputersCount'
                    ExcelData       = [ActiveDirectory]::DomainComputersCount
                    EmptyParagraphsBefore = 1
                }
                DomainServers                                     = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Servers'
                    TocListLevel    = 2
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainServers
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = 'Following client computers are created in <Domain>.'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - DomainComputers'
                    ExcelData       = [ActiveDirectory]::DomainServers
                }
                DomainServersCount                                = [ordered] @{
                    Use             = $true
                    TableData       = [ActiveDirectory]::DomainServersCount
                    TableDesign     = 'ColorfulGridAccent5'
                    TableTitleMerge = $true
                    TableTitleText  = 'Servers Count'
                    Text            = "Following table and chart shows number of servers and their versions"
                    ChartEnable     = $True
                    ChartTitle      = 'Servers Count'
                    ChartData       = [ActiveDirectory]::DomainServersCount
                    ChartKeys       = 'System Name','System Count'
                    ChartValues     = 'System Count'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - DomainServersCount'
                    ExcelData       = [ActiveDirectory]::DomainServersCount
                    EmptyParagraphsBefore = 1
                }
                DomainComputersUnknown                            = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Unknown Computer Objects'
                    TocListLevel    = 2
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainComputersUnknown
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = 'Following client computers are not asisgned to clients or computers in <Domain>.'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - ComputersUnknown'
                    ExcelData       = [ActiveDirectory]::DomainComputersUnknown
                }
                DomainComputersUnknownCount                       = [ordered] @{
                    Use             = $true
                    TableData       = [ActiveDirectory]::DomainComputersUnknownCount
                    TableDesign     = 'ColorfulGridAccent5'
                    TableTitleMerge = $true
                    TableTitleText  = 'Unknown Computers Count'
                    Text            = "Following table and chart shows number of unknown object computers in domain."
                    ExcelExport     = $false
                    ExcelWorkSheet  = '<Domain> - ComputersUnknownCount'
                    ExcelData       = [ActiveDirectory]::DomainComputersUnknownCount
                    EmptyParagraphsBefore = 1
                }
                SectionPasswordQuality                            = [ordered] @{
                    ### Enables section
                    Use              = $true

                    ### Decides how TOC should be visible
                    TocEnable        = $True
                    TocText          = 'Domain Password Quality'
                    TocListLevel     = 1
                    TocListItemType  = [ListItemType]::Numbered
                    TocHeadingType   = [HeadingType]::Heading1

                    ### Text is added before table/list
                    Text             = "This section provides overview about password quality used in <Domain>. One should review if all those potentially" `
                        + " dangerous approaches to password quality should be left as is or addressed in one way or another."
                    TextAlignment    = [Alignment]::Both
                    PageBreaksAfter  = 0
                    PageBreaksBefore = 1
                }
                DomainPasswordClearTextPassword                   = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Passwords with Reversible Encryption'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordClearTextPassword
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'Passwords of these accounts are stored using reversible encryption.'
                    TextNoData          = 'There are no accounts that have passwords stored using reversible encryption.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordClearText'
                    ExcelData           = [ActiveDirectory]::DomainPasswordClearTextPassword
                }
                DomainPasswordLMHash                              = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Passwords with LM Hash'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordLMHash
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'LM-hashes is the oldest password storage used by Windows, dating back to OS/2 system.' `
                        + ' Due to the limited charset allowed, they are fairly easy to crack. Following accounts are affected:'
                    TextNoData          = 'LM-hashes is the oldest password storage used by Windows, dating back to OS/2 system.' `
                        + ' There were no accounts found that use LM Hashes.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordLMHash'
                    ExcelData           = [ActiveDirectory]::DomainPasswordLMHash
                }
                DomainPasswordEmptyPassword                       = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Empty Passwords'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordEmptyPassword
                    TableMaximumColumns = 4
                    TableDesign         = 'ColorfulGridAccent5'
                    Text                = 'Following accounts have no password set:'
                    TextNoData          = 'There are no accounts in <Domain> that have no password set.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordEmpty'
                    ExcelData           = [ActiveDirectory]::DomainPasswordEmptyPassword
                }
                DomainPasswordWeakPassword                        = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Known passwords'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordWeakPassword
                    TableMaximumColumns = 4
                    TableDesign         = 'ColorfulGridAccent5'
                    Text                = "Passwords of these accounts have been found in given dictionary. It's highely recommended to " `
                        + "notify those users and ask them to change their passwords asap!"
                    TextNoData          = 'There were no passwords found that match given dictionary.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordKnown'
                    ExcelData           = [ActiveDirectory]::DomainPasswordWeakPassword
                }
                DomainPasswordDefaultComputerPassword             = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Default Computer Password'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordDefaultComputerPassword
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'These computer objects have their password set to default:'
                    TextNoData          = 'There were no accounts found that match default computer password criteria.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordDefaultComputer'
                    ExcelData           = [ActiveDirectory]::DomainPasswordDefaultComputerPassword
                }
                DomainPasswordPasswordNotRequired                 = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Password Not Required'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordPasswordNotRequired
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'These accounts are not required to have a password. For some accounts it may be perfectly acceptable ' `
                        + ' but for some it may not. Those accounts should be reviewed and accepted or changed to proper security.'
                    TextNoData          = 'There were no accounts found that does not require password.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordNotRequired'
                    ExcelData           = [ActiveDirectory]::DomainPasswordPasswordNotRequired
                }
                DomainPasswordPasswordNeverExpires                = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Non expiring passwords'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordPasswordNeverExpires
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'Following account have do not expire password policy set on them. Those accounts should be reviewed whether ' `
                        + 'allowing them to never expire is good idea and accepted risk.'
                    TextNoData          = 'There are no accounts in <Domain> that never expire.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordNeverExpire'
                    ExcelData           = [ActiveDirectory]::DomainPasswordPasswordNeverExpires
                }
                DomainPasswordAESKeysMissing                      = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - AES Keys Missing'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordAESKeysMissing
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'Following accounts have their Kerberos AES keys missing'
                    TextNoData          = 'There are no accounts that hvae their Kerberos AES keys missing.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordAESMissing'
                    ExcelData           = [ActiveDirectory]::DomainPasswordAESKeysMissing
                }
                DomainPasswordPreAuthNotRequired                  = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Kerberos Pre-Auth Not Required'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordPreAuthNotRequired
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'Kerberos pre-authentication is not required for these accounts'
                    TextNoData          = 'There were no accounts found that do not require pre-authentication.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordPreAuthNot'
                    ExcelData           = [ActiveDirectory]::DomainPasswordPreAuthNotRequired
                }
                DomainPasswordDESEncryptionOnly                   = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Only DES Encryption Allowed'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordDESEncryptionOnly
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'Only DES encryption is allowed to be used with these accounts'
                    TextNoData          = 'There are no account that require only DES encryption.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordDESEncr'
                    ExcelData           = [ActiveDirectory]::DomainPasswordDESEncryptionOnly
                }
                DomainPasswordDelegatableAdmins                   = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Delegatable to Service'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordDelegatableAdmins
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'These accounts are allowed to be delegated to a service:'
                    TextNoData          = 'No accounts were found that are allowed to be delegated to a service.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordDelegatable'
                    ExcelData           = [ActiveDirectory]::DomainPasswordDelegatableAdmins
                }
                DomainPasswordDuplicatePasswordGroups             = [ordered] @{
                    Use                 = $true
                    TocEnable           = $True
                    TocText             = 'Password Quality - Groups of Users With Same Password'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordDuplicatePasswordGroups
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = 'Following groups of users have same passwords:'
                    TextNoData          = 'There are no 2 passwords that are the same in <Domain>.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - PasswordGroups'
                    ExcelData           = [ActiveDirectory]::DomainPasswordDuplicatePasswordGroups
                }
                DomainPasswordHashesWeakPassword                  = [ordered] @{
                    Use                 = $false
                    TocEnable           = $True
                    TocText             = 'Password Quality - Leaked Passwords'
                    TocListLevel        = 2
                    TocListItemType     = 'Numbered'
                    TocHeadingType      = 'Heading2'
                    TableData           = [ActiveDirectory]::DomainPasswordHashesWeakPassword
                    TableDesign         = 'ColorfulGridAccent5'
                    TableMaximumColumns = 4
                    Text                = "Passwords of these accounts have been found in given HASH dictionary (https://haveibeenpwned.com/). It's highely recommended to " `
                        + "notify those users and ask them to change their passwords asap!"
                    TextNoData          = 'There were no passwords found that match in given dictionary.'
                    ExcelExport         = $true
                    ExcelWorkSheet      = '<Domain> - LeakedPasswords'
                    ExcelData           = [ActiveDirectory]::DomainPasswordHashesWeakPassword
                }
                DomainPasswordStats                               = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'Password Quality - Statistics'
                    TocListLevel    = 2
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading2'
                    TableData       = [ActiveDirectory]::DomainPasswordStats
                    TableDesign     = 'ColorfulGridAccent5'
                    TableTitleMerge = $true
                    TableTitleText = 'Password Quality Statistics'
                    Text            = "Following table and chart shows password statistics"
                    ChartEnable     = $True
                    ChartTitle      = 'Password Statistics'
                    ChartData       = [ActiveDirectory]::DomainPasswordStats
                    ChartKeys       = 'Keys'
                    ChartValues     = 'Values'
                    ExcelExport     = $true
                    ExcelWorkSheet  = '<Domain> - PasswordStats'
                    ExcelData       = [ActiveDirectory]::DomainPasswordStats
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
                SectionExcelDomainUsers                           = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Users'
                    ExcelData      = [ActiveDirectory]::DomainUsers
                }
                SectionExcelDomainUsersAll                        = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Users All'
                    ExcelData      = [ActiveDirectory]::DomainUsersAll
                }
                SectionExcelDomainUsersSystemAccounts             = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Users System'
                    ExcelData      = [ActiveDirectory]::DomainUsersSystemAccounts
                }
                SectionExcelDomainUsersNeverExpiring              = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Never Expiring'
                    ExcelData      = [ActiveDirectory]::DomainUsersNeverExpiring
                }
                SectionExcelDomainUsersNeverExpiringInclDisabled  = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Never Expiring incl Disabled'
                    ExcelData      = [ActiveDirectory]::DomainUsersNeverExpiringInclDisabled
                }
                SectionExcelDomainUsersExpiredInclDisabled        = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Expired incl Disabled'
                    ExcelData      = [ActiveDirectory]::DomainUsersExpiredInclDisabled
                }
                SectionExcelDomainUsersExpiredExclDisabled        = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Expired excl Disabled'
                    ExcelData      = [ActiveDirectory]::DomainUsersExpiredExclDisabled
                }
                SectionExcelDomainUsersFullList                   = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Users List Full'
                    ExcelData      = [ActiveDirectory]::DomainUsersFullList
                }
                SectionExcelDomainComputersFullList               = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Computers List'
                    ExcelData      = [ActiveDirectory]::DomainComputersFullList
                }
                SectionExcelDomainGroupsFullList                  = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Groups List'
                    ExcelData      = [ActiveDirectory]::DomainGroupsFullList
                }
                SectionExcelDomainGroupsRest                      = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Groups'
                    ExcelData      = [ActiveDirectory]::DomainGroups
                }
                SectionExcelDomainGroupsSpecial                   = [ordered] @{
                    Use            = $true
                    ExcelExport    = $true
                    ExcelWorkSheet = '<Domain> - Groups Special'
                    ExcelData      = [ActiveDirectory]::DomainGroupsSpecial
                }
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
            }
        }
    }
}

Start-Documentation -Document $Document -Verbose