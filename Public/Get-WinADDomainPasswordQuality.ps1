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
        return
    }
    if (-not (Test-Path -Path $FilePath)) {
        Write-Verbose "Get-WinADDomainPasswordQuality - File path doesn't exists, using hashes set to $UseHashes"
        return
    }
    if ($DomainInformation -eq $null) {
        Write-Verbose "Get-WinADDomainPasswordQuality - No DomainInformation given, no alternative approach either. Terminating password quality check."
        return
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
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.ClearTextPassword -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordLMHash = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.LMHash  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordEmptyPassword = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.EmptyPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordWeakPassword = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.WeakPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordDefaultComputerPassword = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.DefaultComputerPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordPasswordNotRequired = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.PasswordNotRequired  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordPasswordNeverExpires = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.PasswordNeverExpires  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordAESKeysMissing = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.AESKeysMissing  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordPreAuthNotRequired = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.PreAuthNotRequired  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordDESEncryptionOnly = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.DESEncryptionOnly -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
    }
    $Data.DomainPasswordDelegatableAdmins = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.DelegatableAdmins  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersAll
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
                $FullUserInformation = $DomainInformation.DomainUsersAll | Where { $_.SamAccountName -eq $User }
                $FullComputerInformation = $DomainInformation.DomainComputersAll | Where { $_.SamAccountName -eq $User }
                if ($FullUserInformation) {
                    $MergedObject = Merge-Objects -Object1 $FoundUser -Object2 $FullUserInformation
                }
                if ($FullComputerInformation) {
                    $MergedObject = Merge-Objects -Object1 $MergedObject -Object2 $FullComputerInformation
                }
                $Value += $MergedObject

            }
        }
        return $Value
    }
    return $Data
}