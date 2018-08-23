
function Get-WinUsers {
    param(
        $Users,
        $ADCatalog,
        $ADCatalogUsers
    )
    $UserList = @()
    foreach ($U in $Users) {
        $UserList += [ordered] @{
            'Name'                              = $U.Name
            'UserPrincipalName'                 = $U.UserPrincipalName
            'SamAccountName'                    = $U.SamAccountName
            'DisplayName'                       = $U.DisplayName
            'GivenName'                         = $U.GivenName
            'Surname'                           = $U.Surname
            'EmailAddress'                      = $U.EmailAddress
            'PasswordExpired'                   = $U.PasswordExpired
            'PasswordLastSet'                   = $U.PasswordLastSet
            'PasswordNotRequired'               = $U.PasswordNotRequired
            'PasswordNeverExpires'              = $U.PasswordNeverExpires
            'Enabled'                           = $U.Enabled
            'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).Name
            'ManagerEmail'                      = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).EmailAddress
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
        }

    }
    return Format-TransposeTable -Object $UserList
}

function Get-WinUsersFromGroup {
    param(
        $Group,
        $ADCatalog,
        $ADCatalogUsers
    )
    $UserList = @()
    foreach ($U in $Users) {
        $UserList += [ordered] @{
            'Name'                              = $U.Name
            'UserPrincipalName'                 = $U.UserPrincipalName
            'SamAccountName'                    = $U.SamAccountName
            'DisplayName'                       = $U.DisplayName
            'GivenName'                         = $U.GivenName
            'Surname'                           = $U.Surname
            'EmailAddress'                      = $U.EmailAddress
            'PasswordExpired'                   = $U.PasswordExpired
            'PasswordLastSet'                   = $U.PasswordLastSet
            'PasswordNotRequired'               = $U.PasswordNotRequired
            'PasswordNeverExpires'              = $U.PasswordNeverExpires
            'Enabled'                           = $U.Enabled
            'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).Name
            'ManagerEmail'                      = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).EmailAddress
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
        }

    }
    return Format-TransposeTable -Object $UserList
}

function Get-WinGroupMembers {
    param(
        $Groups,
        $Domain,
        $ADCatalog,
        $ADCatalogUsers,
        $Option
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
                    'Name'                              = $Member.Name
                    'SID'                               = $Member.SID.Value
                    'UserPrincipalName'                 = $Object.UserPrincipalName
                    'SamAccountName'                    = $Object.SamAccountName
                    'DisplayName'                       = $Object.DisplayName
                    'GivenName'                         = $Object.GivenName
                    'Surname'                           = $Object.Surname
                    'EmailAddress'                      = $Object.EmailAddress
                    'PasswordExpired'                   = $Object.PasswordExpired
                    'PasswordLastSet'                   = $Object.PasswordLastSet
                    'PasswordNotRequired'               = $Object.PasswordNotRequired
                    'PasswordNeverExpires'              = $Object.PasswordNeverExpires
                    'Enabled'                           = $Object.Enabled
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

                    "Description"                       = $Object.Description
                    "DistinguishedName"                 = $Object.DistinguishedName
                    "EmployeeID"                        = $Object.EmployeeID
                    "EmployeeNumber"                    = $Object.EmployeeNumber
                    "LastBadPasswordAttempt"            = $Object.LastBadPasswordAttempt
                    "LastLogonDate"                     = $Object.LastLogonDate

                    "Created"                           = $Object.Created
                    "Modified"                          = $Object.Modified
                    "Protected"                         = $Object.ProtectedFromAccidentalDeletion
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
                }
            }
        }
        return Format-TransposeTable -Object $GroupMembersDirect
    }
}



Function Get-PrivilegedGroupsMembers {
    [CmdletBinding()]
    param (
        $Domain,
        $DomainSID
    )
    $PrivilegedGroups1 = "$DomainSID-512", "$DomainSID-518", "$DomainSID-519", "$DomainSID-520" # will be only on root domain
    $PrivilegedGroups2 = "S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549", "S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552", "S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573", "S-1-5-32-578", "S-1-5-32-580"

    $SpecialGroups = @()
    foreach ($Group in ($PrivilegedGroups1 + $PrivilegedGroups2)) {
        Write-Verbose "Get-PrivilegedGroupsMembers - Group $Group in $Domain ($DomainSid)"
        try {
            $GroupInfo = Get-AdGroup -Identity $Group -Server $Domain -ErrorAction Stop
            $GroupData = get-adgroupmember -Server $Domain -Identity $group | Sort-Object -Unique
            $GroupDataRecursive = get-adgroupmember -Server $Domain -Identity $group -Recursive:$Recursive | Sort-Object -Unique
            #$GroupDataRecursive | fl *
            #$GroupData.SamAccountName #| Select * -Unique
            #$GroupData | ft -a
            $SpecialGroups += [ordered]@{
                'Group Name'              = $GroupInfo.Name
                'Group Category'          = $GroupInfo.GroupCategory
                'Group Scope'             = $GroupInfo.GroupScope
                'Members Count'           = Get-ObjectCount $GroupData
                'Members Count Recursive' = Get-ObjectCount $GroupDataRecursive
                'Members'                 = $GroupData.SamAccountName
                'Members Recursive'       = $GroupDataRecursive.SamAccountName
            }
        } catch {
            Write-Verbose "Get-PrivilegedGroupsMembers - Error on Group $Group in $Domain ($DomainSid)"
        }
    }
    return $SpecialGroups.ForEach( {[PSCustomObject]$_})
}
function Get-WinCustomGroupInformation {
    [CmdletBinding()]
    param (
        [string] $Domain,
        [string[]] $Groups
    )
    $SpecialGroups = @()
    $AllGroups = Get-ADGroup -Server $Domain -Filter *
    foreach ($Group in $AllGroups) {
        $GroupInfo = Get-AdGroup -Identity $Group -Server $Domain -ErrorAction Stop
        $GroupData = get-adgroupmember -Server $Domain -Identity $group | Sort-Object -Unique
        $GroupDataRecursive = get-adgroupmember -Server $Domain -Identity $group -Recursive:$Recursive | Sort-Object -Unique
        $SpecialGroups += [ordered]@{
            'Group Name'              = $GroupInfo.Name
            'Group Category'          = $GroupInfo.GroupCategory
            'Group Scope'             = $GroupInfo.GroupScope
            'Members Count'           = Get-ObjectCount $GroupData
            'Members Count Recursive' = Get-ObjectCount $GroupDataRecursive
            'Members'                 = $GroupData.SamAccountName
            'Members Recursive'       = $GroupDataRecursive.SamAccountName
            'Group SID'               = $GroupInfo.SID
        }
    }
    return $SpecialGroups.ForEach( {[PSCustomObject]$_}) | Sort-Object 'Group SID'
}
function Get-SpecialGroups {
    param (
        $Groups,
        $Type
    )
    $PrivilegedGroups2 = "S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549", "S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552", "S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573", "S-1-5-32-578", "S-1-5-32-580"
    $PrivilegedGroups1 = "$DomainSID-512", "$DomainSID-518", "$DomainSID-519", "$DomainSID-520" # will be only on root domain

    $Groups = $Groups | Where { ("$($_.'Group SID')").length -eq 12 }

    If ($Type -eq 0) {
        $Yellow = $Groups | Where {$PrivilegedGroups2 -NotContains $_.'Group SID'}
    } elseif ($Type -eq 1) {
        $Yellow = $Groups | Where {$PrivilegedGroups2 -contains $_.'Group SID'}
    } elseif ($Type -eq 2) {
        #$Yellow = $Groups | Where {$PrivilegedGroups1 -c}
    }
    return $Yellow


    #$PrivGroups = @()
    foreach ($Group in $Groups) {
        foreach ($PrivGroup in $PrivilegedGroups2) {
            if ($NotInScope) {
                if ($Group.'Group SID' -ne $PrivGroup) {
                    #  $PrivGroups += $Group
                }
            } else {
                if ($Group.'Group SID' -eq $PrivGroup) {
                    #   $PrivGroups += $Group
                }
            }
        }
    }
    # return $PrivGroups
}