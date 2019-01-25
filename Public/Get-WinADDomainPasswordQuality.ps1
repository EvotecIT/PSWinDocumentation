function Get-WinADDomainPasswordQuality {
    [CmdletBinding()]
    param (
        $DomainInformation,
        $PasswordQualityUsers,
        [string] $FilePath,
        [switch] $UseHashes
    )
    if ([string]::IsNullOrEmpty($FilePath)) {
        Write-Verbose "Get-WinADDomainPasswordQuality - File path not given, using hashes set to $UseHashes"
        return $null
    }
    if (-not (Test-Path -Path $FilePath)) {
        Write-Verbose "Get-WinADDomainPasswordQuality - File path doesn't exists, using hashes set to $UseHashes"
        return $null
    }
    if ($DomainInformation -eq $null) {
        Write-Verbose "Get-WinADDomainPasswordQuality - No DomainInformation given, no alternative approach either. Terminating password quality check."
        return $null
    }
    $Data = [ordered] @{}
    if ($PasswordQualityUsers) {
        $Data.PasswordQualityUsers = $PasswordQualityUsers
    } else {
        $Data.PasswordQualityUsers = Get-ADReplAccount -All -Server $DomainInformation.DomainInformation.DnsRoot -NamingContext $DomainInformation.DomainInformation.DistinguishedName
    }
    $Data.PasswordQuality = Invoke-Command -ScriptBlock {
        if ($UseHashes) {
            $Results = $Data.PasswordQualityUsers | Test-PasswordQuality -WeakPasswordHashesFile $FilePath -IncludeDisabledAccounts
        } else {
            $Results = $Data.PasswordQualityUsers | Test-PasswordQuality -WeakPasswordsFile $FilePath -IncludeDisabledAccounts
        }
        return $Results
    }
    $Data.DomainPasswordClearTextPassword = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList $Data.PasswordQuality.ClearTextPassword -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordClearTextPasswordEnabled = Invoke-Command -ScriptBlock {
        return $Data.DomainPasswordClearTextPassword | Where-Object { $_.Enabled -eq $true }
    }
    $Data.DomainPasswordClearTextPasswordDisabled = Invoke-Command -ScriptBlock {
        return $Data.DomainPasswordClearTextPassword | Where-Object { $_.Enabled -eq $false }
    }
    $Data.DomainPasswordLMHash = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList  $Data.PasswordQuality.LMHash  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordEmptyPassword = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList $Data.PasswordQuality.EmptyPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }


    $Data.DomainPasswordWeakPassword = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList  $Data.PasswordQuality.WeakPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordWeakPasswordEnabled = Invoke-Command -ScriptBlock {
        return $Data.DomainPasswordWeakPassword  | Where-Object { $_.Enabled -eq $true }
    }
    $Data.DomainPasswordWeakPasswordDisabled = Invoke-Command -ScriptBlock {
        return $Data.DomainPasswordWeakPassword  | Where-Object { $_.Enabled -eq $false }
    }
    $Data.DomainPasswordWeakPasswordList = Invoke-Command -ScriptBlock {
        if ($UseHashes) {
            return ''
        } else {
            $Passwords = Get-Content -Path $FilePath
            return $Passwords -join ', '
        }
    }
    $Data.DomainPasswordDefaultComputerPassword = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList  $Data.PasswordQuality.DefaultComputerPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordPasswordNotRequired = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList  $Data.PasswordQuality.PasswordNotRequired  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordPasswordNeverExpires = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList $Data.PasswordQuality.PasswordNeverExpires  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordAESKeysMissing = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList  $Data.PasswordQuality.AESKeysMissing  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordPreAuthNotRequired = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList $Data.PasswordQuality.PreAuthNotRequired  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordDESEncryptionOnly = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList $Data.PasswordQuality.DESEncryptionOnly -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordDelegatableAdmins = Invoke-Command -ScriptBlock {
        $ADAccounts = Get-WinADAccounts -UserNameList $Data.PasswordQuality.DelegatableAdmins  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
        return $ADAccounts | Select-Object 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    $Data.DomainPasswordDuplicatePasswordGroups = Invoke-Command -ScriptBlock {
        $Value = @()
        $DuplicateGroups = $Data.PasswordQuality.DuplicatePasswordGroups.ToArray()
        $Count = 0
        foreach ($DuplicateGroup in $DuplicateGroups) {
            $Count++
            $Name = "Duplicate $Count"
            foreach ($User in $DuplicateGroup) {
                $FoundUser = [pscustomobject] @{
                    'Duplicate Group' = $Name
                    #'Found User'      = $User
                }
                $FullUserInformation = $DomainInformation.DomainUsersAll | Where-Object { $_.SamAccountName -eq $User }
                $FullComputerInformation = $DomainInformation.DomainComputersAll | Where-Object { $_.SamAccountName -eq $User }
                if ($FullUserInformation) {
                    $MergedObject = Merge-Objects -Object1 $FoundUser -Object2 $FullUserInformation
                }
                if ($FullComputerInformation) {
                    $MergedObject = Merge-Objects -Object1 $MergedObject -Object2 $FullComputerInformation
                }
                $Value += $MergedObject

            }
        }
        # Added 'Duplicate Group' to standard output of names - without it, it doesn't make sense
        return $Value | Select-Object 'Duplicate Group', 'Name', 'UserPrincipalName', 'Enabled', 'Password Last Changed', "DaysToExpire", `
            'PasswordExpired', 'PasswordNeverExpires', 'PasswordNotRequired', 'DateExpiry', 'PasswordLastSet', 'SamAccountName', `
            'EmailAddress', 'Display Name', 'Given Name', 'Surname', 'Manager', 'Manager Email', `
            "AccountExpirationDate", "AccountLockoutTime", "AllowReversiblePasswordEncryption", "BadLogonCount", `
            "CannotChangePassword", "CanonicalName", "Description", "DistinguishedName", "EmployeeID", "EmployeeNumber", "LastBadPasswordAttempt", `
            "LastLogonDate", "Created", "Modified", "Protected", "Primary Group", "Member Of", "Domain"
    }
    return $Data
}