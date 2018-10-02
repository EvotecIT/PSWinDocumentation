function Get-WinADDomainPasswordQuality {
    [CmdletBinding()]
    param (
        $DomainInformation,
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
    $Data = [ordered] @{}
    $Data.PasswordQualityUsers = Get-ADReplAccount -All -Server $DomainInformation.DomainInformation.DnsRoot -NamingContext $DomainInformation.DomainInformation.DistinguishedName
    $Data.PasswordQuality = Invoke-Command -ScriptBlock {
        if ($UseHashes) {
            $Results = $Data.PasswordQualityUsers | Test-PasswordQuality -WeakPasswordHashesFile $FilePath -IncludeDisabledAccounts
        } else {
            $Results = $Data.PasswordQualityUsers | Test-PasswordQuality -WeakPasswordsFile $FilePath -IncludeDisabledAccounts
        }
        return $Results
    }

    $Data.DomainPasswordClearTextPassword = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.ClearTextPassword -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordLMHash = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.LMHash  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordEmptyPassword = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.EmptyPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordWeakPassword = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.WeakPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordDefaultComputerPassword = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.DefaultComputerPassword  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordPasswordNotRequired = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.PasswordNotRequired  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordPasswordNeverExpires = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.PasswordNeverExpires  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordAESKeysMissing = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList  $Data.PasswordQuality.AESKeysMissing  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordPreAuthNotRequired = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.PreAuthNotRequired  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordDESEncryptionOnly = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.DESEncryptionOnly -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
    }
    $Data.DomainPasswordDelegatableAdmins = Invoke-Command -ScriptBlock {
        return Get-WinADAccounts -UserNameList $Data.PasswordQuality.DelegatableAdmins  -ADCatalog $DomainInformation.DomainUsersAll, $DomainInformation.DomainComputersFullList
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
                $FullComputerInformation = $DomainInformation.DomainComputersFullList | Where { $_.SamAccountName -eq $User }
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