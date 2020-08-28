function Invoke-ADWord {
    [cmdletBinding()]
    param(
        [string] $FilePath,
        [System.Collections.IDictionary]$DataSetForest
    )
    Documentimo -FilePath $FilePath {
        DocToc -Title 'Table of content'

        DocPageBreak

        DocText {
            "This document provides low-level documentation of Active Directory infrastructure in organization. This document contains general data that has been exported from Active Directory and provides an overview of the whole environment."
        }
        DocNumbering -Text 'General Information - Forest Summary' -Level 0 -Type Numbered -Heading Heading1 {
            DocText {
                "Active Directory has a forest name $($DataSetForest.ForestInformation.Name). Following table contains forest summary with important information:"
            }
            DocTable -DataTable $DataSetForest.ForestInformation -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Forest Summary'
            DocText -LineBreak

            DocText -Text 'Following table contains FSMO servers:'
            DocTable -DataTable $DataSetForest.ForestFSMO -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'FSMO Roles'
            DocText -LineBreak

            DocText -Text 'Following table contains optional forest features:'
            DocTable -DataTable $DataSetForest.ForestOptionalFeatures -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Optional Features'
            DocText -LineBreak

            DocText -Text 'Following UPN suffixes were created in this forest:'
            DocTable -DataTable $DataSetForest.ForestUPNSuffixes -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'UPN Suffixes'
            DocText -LineBreak

            if ($DataSetForest.ForestSPNSuffixes) {
                DocText -Text 'Following SPN suffixes were created in this forest:'
                DocTable -DataTable $DataSetForest.ForestSPNSuffixes -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'SPN Suffixes'
            } else {
                DocText -Text 'No SPN suffixes were created in this forest.'
            }
            DocPageBreak

            DocNumbering -Text 'General Information - Forest Sites' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Forest Sites list can be found below:"
                DocTable -DataTable $DataSetForest.ForestSites1 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
                DocText -LineBreak

                DocText -Text "Forest Sites list can be found below:"
                DocTable -DataTable $DataSetForest.ForestSites2 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
                #  DocText -LineBreak
            }

            DocNumbering -Text 'General Information - Subnets' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Table below contains information regarding relation between Subnets and sites"
                DocTable -DataTable $DataSetForest.ForestSubnets1 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
                DocText -LineBreak

                DocText -Text "Table below contains information regarding relation between Subnets and sites"
                DocTable -DataTable $DataSetForest.ForestSubnets2 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
                # DocText -LineBreak
            }

            DocNumbering -Text 'General Information - Site Links' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Forest Site Links information is available in table below"
                DocTable -DataTable $DataSetForest.ForestSiteLinks -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
                # DocText -LineBreak
            }
        }
        foreach ($Domain in $DataSetForest.FoundDomains.Keys) {
            DocPageBreak

            DocNumbering -Text "General Information - Domain $Domain" -Level 0 -Type Numbered -Heading Heading1 {

                DocNumbering -Text 'General Information - Domain Summary' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following domain exists within forest $($DataSetForest.ForestName):"

                    DocList -Type Bulleted {
                        DocListItem -Level 1 -Text "Domain $($DataSetForest.FoundDomains.$Domain.DomainInformation.DistinguishedName)"
                        DocListItem -Level 2 -Text "Name for fully qualified domain name (FQDN): $($DataSetForest.FoundDomains.$Domain.DomainInformation.DNSRoot)"
                        DocListItem -Level 2 -Text "Name for NetBIOS: $($DataSetForest.FoundDomains.$Domain.DomainInformation.NetBIOSName)"
                    }
                }

                DocNumbering -Text 'General Information - Domain Well Known Folders' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following table has defined well known folders:"
                    DocTable -DataTable $($DataSetForest.FoundDomains.$Domain.DomainWellKnownFolders) -Transpose -OverwriteTitle 'Well known folders' -Design ColorfulGridAccent5 -AutoFit Window
                }

                DocNumbering -Text  'General Information - Domain Controllers' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text 'Following table contains domain controllers'
                    DocTable -DataTable ($DataSetForest.FoundDomains.$Domain.DomainControllers) -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
                    DocText -LineBreak

                    DocText -Text "Following table contains FSMO servers with roles for domain $Domain"
                    DocTable -DataTable ($DataSetForest.FoundDomains.$Domain.DomainFSMO) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "FSMO Roles for $Domain"
                }

                DocNumbering -Text 'General Information - Password Policies' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following table contains password policies for all users within $Domain"
                    DocTable -DataTable $DataSetForest.FoundDomains.$Domain.DomainDefaultPasswordPolicy -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "Default Password Policy for $Domain"
                }

                DocNumbering -Text 'General Information - Fine-grained Password Policies' -Level 1 -Type Numbered -Heading Heading1 {
                    if ($DataSetForest.FoundDomains.$Domain.DomainFineGrainedPolicies) {
                        DocText -Text 'Following table contains Fine-grained password policies'
                        DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainFineGrainedPolicies -Design ColorfulGridAccent5 -AutoFit Window # -OverwriteTitle  "Fine-grained Password Policy for <Domain>"
                    } else {
                        DocText {
                            "Following section should cover fine-grained password policies. " `
                                + "There were no fine-grained password polices defined in $Domain. There was no formal requirement to have them set up."
                        }
                    }
                }
                DocNumbering -Text 'General Information - Group Policies' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following table contains group policies for $Domain"
                    DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainGroupPolicies -Design ColorfulGridAccent5 -AutoFit Window
                }
                DocNumbering -Text 'General Information - Group Policies Details' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following table contains group policies for $Domain"
                    DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainGroupPoliciesDetails -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 6
                }

                DocNumbering -Text 'General Information - DNS A/SRV Records' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following table contains SRV records for Kerberos and LDAP"
                    DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainDNSSRV -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 10
                    DocText -LineBreak

                    DocText -Text "Following table contains A records for Kerberos and LDAP"
                    DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainDNSA -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 10
                }


                DocNumbering -Text 'General Information - Trusts' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following table contains trusts established with domains..."
                    DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainTrusts -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 10
                    DocText -LineBreak
                }

                DocNumbering -Text 'General Information - Organizational Units' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text "Following table contains all OU's created in $Domain"
                    DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainOrganizationalUnits -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 4
                    DocText -LineBreak
                }

                DocNumbering -Text 'General Information - Priviliged Groups' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText -Text 'Following table contains list of priviliged groups and count of the members in it.'
                    DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainGroupsPriviliged -Design ColorfulGridAccent5 -AutoFit Window
                    DocChart -Title 'Priviliged Group Members' -DataTable  $DataSetForest.FoundDomains.$Domain.DomainGroupsPriviliged -Key 'Group Name' -Value 'Member Count'
                }

                DocNumbering -Text "General Information - Domain Users in $Domain" -Level 1 -Type Numbered -Heading Heading1 {
                    DocNumbering -Text 'General Information - Users Count' -Level 2 -Type Numbered -Heading Heading2 {
                        DocText -Text "Following table and chart shows number of users in its categories"
                        DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainUsersCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Users Count'
                        DocChart -Title 'Users Count' -DataTable  $DataSetForest.FoundDomains.$Domain.DomainUsersCount
                    }
                    DocNumbering -Text 'General Information - Domain Administrators' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainAdministratorsRecursive) {
                            DocText -Text 'Following users have highest priviliges and are able to control a lot of Windows resources.'
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainAdministratorsRecursive -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText -Text 'No Domain Administrators users were defined for this domain.'
                        }
                    }
                    DocNumbering -Text  'General Information - Enterprise Administrators' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainEnterpriseAdministratorsRecursive) {
                            DocText -Text 'Following users have highest priviliges across Forest and are able to control a lot of Windows resources.'
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainEnterpriseAdministratorsRecursive -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText -Text 'No Enterprise Administrators users were defined for this domain.'
                        }
                    }
                }

                DocNumbering -Text "General Information - Computer Objects in $Domain" -Level 1 -Type Numbered -Heading Heading1 {


                    DocNumbering -Text 'General Information - Computers' -Level 2 -Type Numbered -Heading Heading2 {
                        DocText -Text "Following table and chart shows number of computers and their versions"
                        DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainComputersCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Computers Count'
                        DocChart -Title 'Computers Count' -DataTable $DataSetForest.FoundDomains.$Domain.DomainComputersCount -Key 'System Name' -Value  'System Count'
                    }
                    DocNumbering -Text 'General Information - Servers' -Level 2 -Type Numbered -Heading Heading2 {
                        DocText -Text "Following table and chart shows number of servers and their versions"
                        DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainServersCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Servers Count'
                        DocChart -Title 'Servers Count' -DataTable $DataSetForest.FoundDomains.$Domain.DomainServersCount -Key 'System Name' -Value  'System Count'
                    }
                    DocNumbering -Text 'General Information - Unknown Computers' -Level 2 -Type Numbered -Heading Heading2 {
                        DocText -Text "Following table and chart shows number of unknown object computers in domain."
                        DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainComputersUnknownCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Unknown Computers Count'
                        DocChart -Title 'Unknown Computers Count' -DataTable $DataSetForest.FoundDomains.$Domain.DomainComputersUnknownCount -Key 'System Name' -Value  'System Count'
                    }
                }

                DocNumbering -Text 'Domain Password Quality' -Level 1 -Type Numbered -Heading Heading1 {
                    DocText {
                        "This section provides overview about password quality used in $Domain. One should review if all those potentially" `
                            + " dangerous approaches to password quality should be left as is or addressed in one way or another."
                    }

                    DocNumbering -Text 'Password Quality - Passwords with Reversible Encryption' -Level 2 -Type Numbered -Heading Heading2 {
                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordClearTextPassword) {
                            DocText -Text 'Passwords of these accounts are stored using reversible encryption.'
                            DocTable -DataTable $DataSetForest.FoundDomains.$Domain.DomainPasswordClearTextPassword -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText -Text 'There are no accounts that have passwords stored using reversible encryption.'
                        }
                    }
                    DocNumbering -Text 'Password Quality - Passwords with LM Hash' -Level 2 -Type Numbered -Heading Heading2 {
                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordLMHash) {
                            DocText {
                                'LM-hashes is the oldest password storage used by Windows, dating back to OS/2 system.'
                                'Due to the limited charset allowed, they are fairly easy to crack. Following accounts are affected:'
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordLMHash -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                'LM-hashes is the oldest password storage used by Windows, dating back to OS/2 system.' `
                                    + ' There were no accounts found that use LM Hashes.'
                            }
                        }
                    }
                    DocNumbering -Text 'Password Quality - Empty Passwords' -Level 2 -Type Numbered -Heading Heading2 {
                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordEmptyPassword) {
                            DocText -Text  'Following accounts have no password set:'
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordEmptyPassword -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText -Text "There are no accounts in $Domain that have no password set."
                        }
                    }
                    DocNumbering -Text 'Password Quality - Known passwords' -Level 2 -Type Numbered -Heading Heading2 {
                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordWeakPassword) {
                            DocText {
                                "Passwords of these accounts have been found in given dictionary. It's highly recommended to " `
                                    + "notify those users and ask them to change their passwords asap! Following is list of passwords used for comparision:"
                            }
                            DocList {
                                foreach ($_ in $DataSetForest.FoundDomains.$Domain.DomainPasswordWeakPasswordList) {
                                    DocListItem -Level 1 -Text $_
                                }
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordWeakPassword -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText -Text 'Domain was tested for the use of known passwords from the list below:'
                            DocList {
                                foreach ($_ in $DataSetForest.FoundDomains.$Domain.DomainPasswordWeakPasswordList) {
                                    DocListItem -Level 1 -Text $_
                                }
                            }
                            DocText -Text 'There were no accounts found that match given dictionary.'
                        }
                    }
                    DocNumbering -Text 'Password Quality - Default Computer Password' -Level 2 -Type Numbered -Heading Heading2 {
                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordDefaultComputerPassword) {
                            DocText -Text 'These computer objects have their password set to default:'
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordDefaultComputerPassword -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText -Text 'There were no accounts found that match default computer password criteria.'
                        }
                    }
                    DocNumbering -Text 'Password Quality - Password Not Required' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordPasswordNotRequired) {
                            DocText {
                                'These accounts are not required to have a password. For some accounts it may be perfectly acceptable ' `
                                    + ' but for some it may not. Those accounts should be reviewed and accepted or changed to proper security.'
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordPasswordNotRequired -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                'There were no accounts found that does not require password.'
                            }
                        }
                    }
                    DocNumbering -Text  'Password Quality - Non expiring passwords' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordPasswordNeverExpires) {
                            DocText {
                                'Following account have do not expire password policy set on them. Those accounts should be reviewed whether ' `
                                    + 'allowing them to never expire is good idea and accepted risk.'
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordPasswordNeverExpires -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                "There are no accounts in $Domain that never expire."
                            }
                        }
                    }
                    DocNumbering -Text 'Password Quality - AES Keys Missing' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordAESKeysMissing) {
                            DocText {
                                'Following accounts have their Kerberos AES keys missing'
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordAESKeysMissing -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                'There are no accounts that hvae their Kerberos AES keys missing.'
                            }
                        }
                    }
                    DocNumbering -Text 'Password Quality - Kerberos Pre-Auth Not Required' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordPreAuthNotRequired) {
                            DocText {
                                'Kerberos pre-authentication is not required for these accounts'
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordPreAuthNotRequired -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                'There were no accounts found that do not require pre-authentication.'
                            }
                        }
                    }
                    DocNumbering -Text 'Password Quality - Only DES Encryption Allowed' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordDESEncryptionOnly) {
                            DocText {
                                'Only DES encryption is allowed to be used with these accounts'
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordDESEncryptionOnly -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                'There are no account that require only DES encryption.'
                            }
                        }
                    }
                    DocNumbering -Text 'Password Quality - Delegatable to Service' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordDelegatableAdmins) {
                            DocText {
                                'These accounts are allowed to be delegated to a service:'
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordDelegatableAdmins -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                'No accounts were found that are allowed to be delegated to a service.'
                            }
                        }
                    }
                    DocNumbering -Text 'Password Quality - Groups of Users With Same Password' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordDuplicatePasswordGroups) {
                            DocText {
                                'Following groups of users have same passwords:'
                            }
                            DocTable -DataTable $DataSetForest.FoundDomains.$Domain.DomainPasswordDuplicatePasswordGroups -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                "There are no 2 passwords that are the same in $Domain."
                            }
                        }
                    }
                    DocNumbering -Text 'Password Quality - Leaked Passwords' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordHashesWeakPassword) {
                            DocText {
                                "Passwords of these accounts have been found in given HASH dictionary (https://haveibeenpwned.com/). It's highly recommended to " `
                                    + "notify those users and ask them to change their passwords asap!"
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordHashesWeakPassword -Design ColorfulGridAccent5 -AutoFit Window
                        } else {
                            DocText {
                                'There were no passwords found that match in given dictionary.'
                            }
                        }
                    }
                    DocNumbering -Text  'Password Quality - Statistics' -Level 2 -Type Numbered -Heading Heading2 {

                        if ($DataSetForest.FoundDomains.$Domain.DomainPasswordStats) {
                            DocText {
                                "Following table and chart shows password statistics"
                            }
                            DocTable -DataTable  $DataSetForest.FoundDomains.$Domain.DomainPasswordStats -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Password Quality - Statistics'
                        } else {
                            DocText {
                                'There were no passwords found that match in given dictionary.'
                            }
                        }
                        DocChart -Title 'Password Statistics' -DataTable $DataSetForest.FoundDomains.$Domain.DomainPasswordStats # Hashtables don't require Key/Value pair
                    }
                }
            }
        }
    } -Open
}