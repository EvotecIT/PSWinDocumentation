function Get-WinUsers {
    [CmdletBinding()]
    param(
        [System.Object[]] $Users,
        [System.Object[]] $ADCatalog,
        [System.Object[]] $ADCatalogUsers,
        [string] $Domain
    )
    [DateTime] $CurrentDate = Get-Date # [DateTime]::Today
    $UserList = foreach ($U in $Users) {
        [PsCustomObject][Ordered] @{
            'Name'                              = $U.Name
            'UserPrincipalName'                 = $U.UserPrincipalName
            'SamAccountName'                    = $U.SamAccountName
            'Display Name'                      = $U.DisplayName
            'Given Name'                        = $U.GivenName
            'Surname'                           = $U.Surname
            'EmailAddress'                      = $U.EmailAddress
            'PasswordExpired'                   = $U.PasswordExpired
            'PasswordLastSet'                   = $U.PasswordLastSet
            'Password Last Changed'             = if ($U.PasswordLastSet -ne $Null) { "$(-$($U.PasswordLastSet - $CurrentDate).Days) days" } else { 'N/A'}
            'PasswordNotRequired'               = $U.PasswordNotRequired
            'PasswordNeverExpires'              = $U.PasswordNeverExpires
            'Enabled'                           = $U.Enabled
            'Manager'                           = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).Name
            'Manager Email'                     = (Get-ADObjectFromDistingusishedName -ADCatalog $ADCatalogUsers -DistinguishedName $U.Manager).EmailAddress
            'DateExpiry'                        = Convert-ToDateTime -Timestring $($U."msDS-UserPasswordExpiryTimeComputed") -Verbose
            "DaysToExpire"                      = (Convert-TimeToDays -StartTime $CurrentDate -EndTime (Convert-ToDateTime -Timestring $($U."msDS-UserPasswordExpiryTimeComputed")))
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
    return $UserList
}

<# List of fields
'Name', 'UserPrincipalName', 'SamAccountName', 'Enabled', 'PasswordLastSet,'Password Last Changed', 'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired',
'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email',
'DateExpiry', "DaysToExpire", "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount",
"CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt",
"LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
#>