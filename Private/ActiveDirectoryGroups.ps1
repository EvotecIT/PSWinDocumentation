function Get-WinUsers {
    param(
        [System.Object[]] $Users,
        [System.Object[]] $ADCatalog,
        [System.Object[]] $ADCatalogUsers,
        [string] $Domain
    )
    $UserList = @()
    foreach ($U in $Users) {
        $UserList += [ordered] @{
            'Name'                              = $U.Name
            'UserPrincipalName'                 = $U.UserPrincipalName
            'SamAccountName'                    = $U.SamAccountName
            'Display Name'                      = $U.DisplayName
            'Given Name'                        = $U.GivenName
            'Surname'                           = $U.Surname
            'EmailAddress'                      = $U.EmailAddress
            'PasswordExpired'                   = $U.PasswordExpired
            'PasswordLastSet'                   = $U.PasswordLastSet
            'PasswordNotRequired'               = $U.PasswordNotRequired
            'PasswordNeverExpires'              = $U.PasswordNeverExpires
            'Enabled'                           = $U.Enabled
            'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).Name
            'Manager Email'                     = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).EmailAddress
            'DateExpiry'                        = Convert-ToDateTime -Timestring $($U."msDS-UserPasswordExpiryTimeComputed") -Verbose
            "DaysToExpire"                      = (Convert-TimeToDays -StartTime GET-DATE -EndTime (Convert-ToDateTime -Timestring $($U."msDS-UserPasswordExpiryTimeComputed")))
            "AccountExpirationDate"             = $U.AccountExpirationDate
            "AccountLockoutTime"                = $U.AccountLockoutTime
            "AllowReversiblePasswordEncryption" = $U.AllowReversiblePasswordEncryption
            "BadLogonCount"                     = $U.BadLogonCount
            "CannotChangePassword"              = $U.CannotChangePassword
            "CanonicalName"                     = $U.CanonicalName

            "Description"                       = $U.Description
            "DistinguishedName"                 = $U.DistinguishedName
            "EmployeeID"                        = $U.EmployeeID
            "EmployeeNumber"                    = $U.EmployeeNumber
            "LastBadPasswordAttempt"            = $U.LastBadPasswordAttempt
            "LastLogonDate"                     = $U.LastLogonDate

            "Created"                           = $U.Created
            "Modified"                          = $U.Modified
            "Protected"                         = $U.ProtectedFromAccidentalDeletion

            "Primary Group"                     = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalog -DistinguishedName $U.PrimaryGroup -Type 'SamAccountName')
            "Member Of"                         = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalog -DistinguishedName $U.MemberOf -Type 'SamAccountName' -Splitter ', ')
            "Domain"                            = $Domain
        }

    }
    return Format-TransposeTable -Object $UserList
}

function Get-WinGroups {
    param (
        [System.Object[]] $Groups,
        [System.Object[]] $Users,
        [string] $Domain
    )
    $ReturnGroups = @()
    foreach ($Group in $Groups) {
        $User = $Users | Where { $_.DistinguishedName -eq $Group.ManagedBy }
        $ReturnGroups += [ordered] @{
            'Group Name'            = $Group.Name
            #'Group Display Name' = $Group.DisplayName
            'Group Category'        = $Group.GroupCategory
            'Group Scope'           = $Group.GroupScope
            'Group SID'             = $Group.SID.Value
            'High Privileged Group' = if ($Group.adminCount -eq 1) { $True } else { $False }
            'Member Count'          = $Group.Members.Count
            'MemberOf Count'        = $Group.MemberOf.Count
            'Manager'               = $User.Name
            'Manager Email'         = $User.EmailAddress
            'Group Members'         = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList, $Data.DomainComputersFullList, $Data.DomainGroupsFullList -DistinguishedName $Group.Members -Type 'SamAccountName')
            'Group Members DN'      = $Group.Members
            "Domain"                = $Domain
        }
    }
    return Format-TransposeTable -Object $ReturnGroups
}

function Get-WinGroupMembers {
    param(
        [System.Object[]] $Groups,
        [string] $Domain,
        [System.Object[]] $ADCatalog,
        [System.Object[]] $ADCatalogUsers,
        [ValidateSet("Recursive", "Standard")][String] $Option
    )
    if ($Option -eq 'Recursive') {
        $GroupMembersRecursive = @()
        foreach ($Group in $Groups) {
            $GroupMembership = Get-ADGroupMember -Server $Domain -Identity $Group.'Group SID' -Recursive
            foreach ($Member in $GroupMembership) {
                $Object = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalog -DistinguishedName $Member.DistinguishedName)
                $GroupMembersRecursive += [ordered] @{
                    'Group Name'                        = $Group.'Group Name'
                    'Group SID'                         = $Group.'Group SID'
                    'Group Category'                    = $Group.'Group Category'
                    'Group Scope'                       = $Group.'Group Scope'
                    'High Privileged Group'             = if ($Group.adminCount -eq 1) { $True } else { $False }
                    'Display Name'                      = $Object.DisplayName
                    'Name'                              = $Member.Name
                    'User Principal Name'               = $Object.UserPrincipalName
                    'Sam Account Name'                  = $Object.SamAccountName
                    'Email Address'                     = $Object.EmailAddress
                    'PasswordExpired'                   = $Object.PasswordExpired
                    'PasswordLastSet'                   = $Object.PasswordLastSet
                    'PasswordNotRequired'               = $Object.PasswordNotRequired
                    'PasswordNeverExpires'              = $Object.PasswordNeverExpires
                    'Enabled'                           = $Object.Enabled
                    'SID'                               = $Member.SID.Value
                    'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $Object.Manager).Name
                    'ManagerEmail'                      = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $Object.Manager).EmailAddress
                    'DateExpiry'                        = Convert-ToDateTime -Timestring $($Object."msDS-UserPasswordExpiryTimeComputed") # -Verbose
                    "DaysToExpire"                      = (Convert-TimeToDays -StartTime GET-DATE -EndTime (Convert-ToDateTime -Timestring $($Object."msDS-UserPasswordExpiryTimeComputed")))
                    "AccountExpirationDate"             = $Object.AccountExpirationDate
                    "AccountLockoutTime"                = $Object.AccountLockoutTime
                    "AllowReversiblePasswordEncryption" = $Object.AllowReversiblePasswordEncryption
                    "BadLogonCount"                     = $Object.BadLogonCount
                    "CannotChangePassword"              = $Object.CannotChangePassword
                    "CanonicalName"                     = $Object.CanonicalName

                    'Given Name'                        = $Object.GivenName
                    'Surname'                           = $Object.Surname

                    "Description"                       = $Object.Description
                    "DistinguishedName"                 = $Object.DistinguishedName
                    "EmployeeID"                        = $Object.EmployeeID
                    "EmployeeNumber"                    = $Object.EmployeeNumber
                    "LastBadPasswordAttempt"            = $Object.LastBadPasswordAttempt
                    "LastLogonDate"                     = $Object.LastLogonDate

                    "Created"                           = $Object.Created
                    "Modified"                          = $Object.Modified
                    "Protected"                         = $Object.ProtectedFromAccidentalDeletion
                    "Domain"                            = $Domain
                }
                # $Member
            }
        }
        return Format-TransposeTable -Object $GroupMembersRecursive
    }
    if ($Option -eq 'Standard') {
        $GroupMembersDirect = @()
        foreach ($Group in $Groups) {
            foreach ($Member in $Group.'Group Members DN') {
                $Object = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalog -DistinguishedName $Member)
                $GroupMembersDirect += [ordered] @{
                    'Group Name'                        = $Group.'Group Name'
                    'Group SID'                         = $Group.'Group SID'
                    'Group Category'                    = $Group.'Group Category'
                    'Group Scope'                       = $Group.'Group Scope'
                    'DisplayName'                       = $Object.DisplayName
                    'High Privileged Group'             = if ($Group.adminCount -eq 1) { $True } else { $False }
                    'UserPrincipalName'                 = $Object.UserPrincipalName
                    'SamAccountName'                    = $Object.SamAccountName
                    'EmailAddress'                      = $Object.EmailAddress
                    'PasswordExpired'                   = $Object.PasswordExpired
                    'PasswordLastSet'                   = $Object.PasswordLastSet
                    'PasswordNotRequired'               = $Object.PasswordNotRequired
                    'PasswordNeverExpires'              = $Object.PasswordNeverExpires
                    'Enabled'                           = $Object.Enabled
                    'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $Object.Manager).Name
                    'ManagerEmail'                      = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $Object.Manager).EmailAddress
                    'DateExpiry'                        = Convert-ToDateTime -Timestring $($Object."msDS-UserPasswordExpiryTimeComputed") #-Verbose
                    "DaysToExpire"                      = (Convert-TimeToDays -StartTime GET-DATE -EndTime (Convert-ToDateTime -Timestring $($Object."msDS-UserPasswordExpiryTimeComputed")))
                    "AccountExpirationDate"             = $Object.AccountExpirationDate
                    "AccountLockoutTime"                = $Object.AccountLockoutTime
                    "AllowReversiblePasswordEncryption" = $Object.AllowReversiblePasswordEncryption
                    "BadLogonCount"                     = $Object.BadLogonCount
                    "CannotChangePassword"              = $Object.CannotChangePassword
                    "CanonicalName"                     = $Object.CanonicalName

                    "Description"                       = $Object.Description
                    "DistinguishedName"                 = $Object.DistinguishedName
                    "EmployeeID"                        = $Object.EmployeeID
                    "EmployeeNumber"                    = $Object.EmployeeNumber
                    "LastBadPasswordAttempt"            = $Object.LastBadPasswordAttempt
                    "LastLogonDate"                     = $Object.LastLogonDate

                    'Name'                              = $Object.Name
                    'SID'                               = $Object.SID.Value
                    'GivenName'                         = $Object.GivenName
                    'Surname'                           = $Object.Surname

                    "Created"                           = $Object.Created
                    "Modified"                          = $Object.Modified
                    "Protected"                         = $Object.ProtectedFromAccidentalDeletion
                    "Domain"                            = $Domain
                }
            }
        }
        return Format-TransposeTable -Object $GroupMembersDirect
    }
}