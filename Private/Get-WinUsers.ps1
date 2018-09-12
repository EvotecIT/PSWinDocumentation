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