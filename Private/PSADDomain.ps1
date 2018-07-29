function Get-ActiveDirectoryData {
    param(
        [string] $Domain
    )
    $ADSnapshot = @{}
    $ADSnapshot.RootDSE = $(Get-ADRootDSE -Server $Domain)
    $ADSnapshot.ForestInformation = $(Get-ADForest -Server $Domain)
    $ADSnapshot.DomainInformation = $(Get-ADDomain -Server $Domain)
    $ADSnapshot.DomainControllers = $(Get-ADDomainController -Server $Domain -Filter * )
    $ADSnapshot.DomainTrusts = (Get-ADTrust -Server $Domain -Filter * )
    $ADSnapshot.DefaultPassWordPoLicy = $(Get-ADDefaultDomainPasswordPolicy -Server $Domain)
    $ADSnapshot.AuthenticationPolicies = $(Get-ADAuthenticationPolicy -Server $Domain -LDAPFilter '(name=AuthenticationPolicy*)')
    $ADSnapshot.AuthenticationPolicySilos = $(Get-ADAuthenticationPolicySilo -Server $Domain -Filter 'Name -like "*AuthenticationPolicySilo*"')
    $ADSnapshot.CentralAccessPolicies = $(Get-ADCentralAccessPolicy -Server $Domain -Filter * )
    $ADSnapshot.CentralAccessRules = $(Get-ADCentralAccessRule -Server $Domain -Filter * )
    $ADSnapshot.ClaimTransformPolicies = $(Get-ADClaimTransformPolicy -Server $Domain -Filter * )
    $ADSnapshot.ClaimTypes = $(Get-ADClaimType -Server $Domain -Filter * )
    $ADSnapshot.DomainAdministrators = $( Get-ADGroup -Identity $('{0}-512' -f (Get-ADDomain).domainSID) | Get-ADGroupMember -Recursive | Get-ADUser)
    $ADSnapshot.OrganizationalUnits = $(Get-ADOrganizationalUnit -Server $Domain -Properties * -Filter * )
    $ADSnapshot.Sites = $(Get-ADReplicationSite -Server $Domain -Filter * )
    $ADSnapshot.Subnets = $(Get-ADReplicationSubnet -Server $Domain -Filter * )
    $ADSnapshot.SiteLinks = $(Get-ADReplicationSiteLink -Server $Domain -Filter * )
    $ADSnapshot.LDAPDNS = $(Resolve-DnsName -Name "_ldap._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.KerberosDNS = $(Resolve-DnsName -Name "_kerberos._tcp.$((Get-ADDomain).DNSRoot)" -Type srv)
    $ADSnapshot.GroupPolicies = $(Get-GPO -Domain $Domain -All) # DisplayName, Owner, DomainName, CreationTime, ModificationTime, GpoStatus, WmiFilter, Description # Id, UserVersion, ComputerVersion
    $ADSnapshot.Users = Invoke-Command -ScriptBlock {
        param(
            $Domain
        )
        function Find-AllUsers {
            param (
                $Domain
            )
            $users = Get-ADUser -Server $Domain -ResultPageSize 5000000 -filter * -Properties Name, Manager, DisplayName, GivenName, Surname, SamAccountName, EmailAddress, msDS-UserPasswordExpiryTimeComputed, PasswordExpired, PasswordLastSet, PasswordNotRequired, PasswordNeverExpires
            $users = $users | Select-Object Name, UserPrincipalName, SamAccountName, DisplayName, GivenName, Surname, EmailAddress, PasswordExpired, PasswordLastSet, PasswordNotRequired, PasswordNeverExpires, Enabled,
            @{Name = "Manager"; Expression = { (Get-ADUser -Server $Domain $_.Manager).Name }},
            @{Name = "ManagerEmail"; Expression = { (Get-ADUser -Server $Domain -Properties Mail $_.Manager).Mail  }},
            @{Name = "DateExpiry"; Expression = { ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")) }},
            @{Name = "DaysToExpire"; Expression = { (NEW-TIMESPAN -Start (GET-DATE) -End ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed"))).Days }}
            return $users
        }
        $Users = Find-AllUsers -Domain $Domain
        return [ordered] @{
            Users                          = $Users
            UsersAll                       = $Users | Where { $_.PasswordNotRequired -eq $False } | Select Name, SamAccountName, UserPrincipalName, Enabled
            UsersSystemAccounts            = $Users | Where { $_.PasswordNotRequired -eq $true } | Select Name, SamAccountName, UserPrincipalName, Enabled
            UsersNeverExpiring             = $Users | Where { $_.PasswordNeverExpires -eq $true -and $_.Enabled -eq $true -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
            UsersNeverExpiringInclDisabled = $Users | Where { $_.PasswordNeverExpires -eq $true -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
            UsersExpiredInclDisabled       = $Users | Where { $_.PasswordNeverExpires -eq $false -and $_.DaysToExpire -le 0 -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
            UsersExpiredExclDisabled       = $Users | Where { $_.PasswordNeverExpires -eq $false -and $_.DaysToExpire -le 0 -and $_.Enabled -eq $true -and $_.PasswordNotRequired -eq $false } | Select Name, SamAccountName, UserPrincipalName, Enabled
        }
    } -ArgumentList $Domain
    $ADSnapshot.UsersCount = [ordered] @{
        'Users Count Incl. System'            = Get-ObjectCount -Object $ADSnapshot.Users.Users
        'Users Count'                         = Get-ObjectCount -Object $ADSnapshot.Users.UsersAll
        'Users Expired'                       = Get-ObjectCount -Object $ADSnapshot.Users.UsersExpiredExclDisabled
        'Users Expired Incl. Disabled'        = Get-ObjectCount -Object $ADSnapshot.Users.UsersExpiredInclDisabled
        'Users Never Expiring'                = Get-ObjectCount -Object $ADSnapshot.Users.UsersNeverExpiring
        'Users Never Expiring Incl. Disabled' = Get-ObjectCount -Object $ADSnapshot.Users.UsersNeverExpiringInclDisabled
        'Users System Accounts'               = Get-ObjectCount -Object $ADSnapshot.Users.UsersSystemAccounts
    }

    return $ADSnapshot
}

function Get-WinDomainInformation {
    param (
        [string] $Domain
    )
    $ADSnapshot = Get-ActiveDirectoryData -Domain $Domain
    $Data = @{}
    $Data.ADSnapshot = $ADSnapshot
    $Data.RootDSE = $ADSnapshot.RootDSE
    $Data.DomainInformation = $ADSnapshot.DomainInformation
    $Data.FSMO = [ordered] @{
        #'Domain Naming Master'  = $ADSnapshot.ForestInformation.DomainNamingMaster
        #'Schema Master'         = $ADSnapshot.ForestInformation.SchemaMaster
        'PDC Emulator'          = $ADSnapshot.DomainInformation.PDCEmulator
        'RID Master'            = $ADSnapshot.DomainInformation.RIDMaster
        'Infrastructure Master' = $ADSnapshot.DomainInformation.InfrastructureMaster
    }
    $Data.GroupPolicies = Invoke-Command -ScriptBlock {
        $GroupPolicies = @()
        foreach ($gpo in $ADSnapshot.GroupPolicies) {
            $GroupPolicy = [ordered] @{
                'Display Name'      = $gpo.DisplayName
                'Gpo Status'        = $gpo.GPOStatus
                'Creation Time'     = $gpo.CreationTime
                'Modification Time' = $gpo.ModificationTime
                'Description'       = $gpo.Description
                'Wmi Filter'        = $gpo.WmiFilter
            }
            $GroupPolicies += $GroupPolicy
        }
        return $GroupPolicies.ForEach( {[PSCustomObject]$_})
    }
    $Data.DefaultPassWordPoLicy = [ordered] @{
        'Complexity Enabled'            = $ADSnapshot.DefaultPassWordPoLicy.ComplexityEnabled
        #'Distinguished Name'            = $ADSnapshot.DefaultPassWordPoLicy.DistinguishedName
        'Lockout Duration'              = $ADSnapshot.DefaultPassWordPoLicy.LockoutDuration
        'Lockout Observation Window'    = $ADSnapshot.DefaultPassWordPoLicy.LockoutObservationWindow
        'Lockout Threshold'             = $ADSnapshot.DefaultPassWordPoLicy.LockoutThreshold
        'Max Password Age'              = $ADSnapshot.DefaultPassWordPoLicy.MaxPasswordAge
        'Min Password Age'              = $ADSnapshot.DefaultPassWordPoLicy.MinPasswordAge
        'Min Password Length'           = $ADSnapshot.DefaultPassWordPoLicy.MinPasswordAge
        'Password History Count'        = $ADSnapshot.DefaultPassWordPoLicy.PasswordHistoryCount
        'Reversible Encryption Enabled' = $ADSnapshot.DefaultPassWordPoLicy.ReversibleEncryptionEnabled
    }
    $Data.PriviligedGroupMembers = Get-PrivilegedGroupsMembers -Domain $Data.DomainInformation.DNSRoot -DomainSID $Data.DomainInformation.DomainSid
    $Data.OrganizationalUnits = $ADSnapshot.OrganizationalUnits | Select-Object Name, CanonicalName, Created | Sort-Object CanonicalName
    $Data.DomainAdministrators = $ADSnapshot.DomainAdministrators | Select-Object Name, SamAccountName, UserPrincipalName, Enabled
    return $Data
}