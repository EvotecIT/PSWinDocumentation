function Get-WinO365Azure {
    [CmdletBinding()]
    param(
        $TypesRequired
    )
    $Data = [ordered] @{}
    if ($TypesRequired -eq $null) {
        Write-Verbose 'Get-WinO365Azure - TypesRequired is null. Getting all Office 365 types.'
        $TypesRequired = Get-Types -Types ([O365])  # Gets all types
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureLicensing, [O365]::O365AzureLicensing)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureLicensing"
        $Data.O365UAzureLicensing = Get-MsolAccountSku
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureLicensing)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureLicensing (prepared data)"
        $Data.O365AzureLicensing = Invoke-Command -ScriptBlock {
            $Licenses = @()
            foreach ($License in $Data.O365UAzureLicensing) {
                $LicensesTotal = $License.ActiveUnits + $License.WarningUnits
                $LicensesUsed = $License.ConsumedUnits
                $LicensesLeft = $LicensesTotal - $LicensesUsed
                $LicenseName = $($Global:O365SKU).Item("$($License.SkuPartNumber)")
                if ($LicenseName -eq $null) { $LicenseName = $License.SkuPartNumber}

                $Licenses += [PSCustomObject] @{
                    Name                 = $LicenseName
                    'Licenses Total'     = $LicensesTotal
                    'Licenses Used'      = $LicensesUsed
                    'Licenses Left'      = $LicensesLeft
                    'Licenses Active'    = $License.ActiveUnits
                    'Licenses Trial'     = $License.WarningUnits
                    'Licenses LockedOut' = $License.LockedOutUnits
                    'Licenses Suspended' = $License.SuspendedUnits
                    'Percent Used'       = ($LicensesUsed / $LicensesTotal).ToString("P")
                    'Percent Left'       = ($LicensesLeft / $LicensesTotal).ToString("P")
                    SKU                  = $License.SkuPartNumber
                    SKUAccount           = $License.AccountSkuId
                    SKUID                = $License.SkuId
                }
            }
            return $Licenses | Sort-Object Name
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureTenantDomains, [O365]::O365AzureTenantDomains)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureTenantDomains"
        $Data.O365UAzureTenantDomains = Get-MsolDomain
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureSubscription, [O365]::O365AzureSubscription)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureSubscription"
        $Data.O365UAzureSubscription = Get-MsolSubscription
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureSubscription)) {
        $Data.O365AzureSubscription = Invoke-Command -ScriptBlock {
            $Licenses = @()
            foreach ($Subscription in $Data.O365UAzureSubscription) {
                $LicenseName = $($Global:O365SKU).Item("$($Subscription.SkuPartNumber)")
                if ($LicenseName -eq $null) { $LicenseName = $Subscription.SkuPartNumber}
                foreach ($Plan in $Subscription.ServiceStatus) {
                    $ServicePlanName = $($Global:O365SKU).Item("$($Plan.ServicePlan.ServiceName)")
                    if ($ServicePlanName -eq $null) { $ServicePlanName = $Plan.ServicePlan.ServiceName}

                    $Licenses += [PSCustomObject] @{
                        'Licenses Name'       = $LicenseName
                        'Licenses SKU'        = $Subscription.SkuPartNumber
                        'Service Plan Name'   = $ServicePlanName
                        'Service Plan SKU'    = $Plan.ServicePlan.ServiceName
                        'Service Plan ID'     = $Plan.ServicePlan.ServicePlanId
                        'Service Plan Type'   = $Plan.ServicePlan.ServiceType
                        'Service Plan Class'  = $Plan.ServicePlan.TargetClass
                        'Service Plan Status' = $Plan.ProvisioningStatus
                        'Licenses Total'      = $Subscription.TotalLicenses
                        'Licenses Status'     = $Subscription.Status
                        'Licenses SKUID'      = $Subscription.SkuId
                        'Licenses Are Trial'  = $Subscription.IsTrial
                        'Licenses Created'    = $Subscription.DateCreated
                        'Next Lifecycle Date' = $Subscription.NextLifecycleDate
                        'ObjectID'            = $Subscription.ObjectId
                        'Ocp SubscriptionID'  = $Subscription.OcpSubscriptionId
                    }
                }
            }
            return $Licenses | Sort-Object 'Licenses Name'
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [O365]::O365UAzureADUsers,
            [O365]::O365AzureADUsersMFA,
            [O365]::O365AzureADUsersStatisticsByCountry,
            [O365]::O365AzureADUsersStatisticsByCity,
            [O365]::O365AzureADUsersStatisticsByCountryCity
        )) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureADUsers"
        $Data.O365UAzureADUsers = Get-MsolUser -All
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureADUsersDeleted)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureADUsersDeleted"
        $Data.O365UAzureADUsersDeleted = Get-MsolUser -ReturnDeletedUsers
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureADGroups, [O365]::O365AzureADGroupMembersUser)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureADGroups"
        $Data.O365UAzureADGroups = Get-MsolGroup -All
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureADGroupMembers, [O365]::O365AzureADGroupMembersUser)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureADGroupMembers"
        $Data.O365UAzureADGroupMembers = Invoke-Command -ScriptBlock {
            $GroupMembers = @()
            foreach ($Group in $Data.Groups) {
                $Object = Get-MsolGroupMember -GroupObjectId $Group.ObjectId -All
                $Object | Add-Member -MemberType NoteProperty -Name "GroupObjectID" -Value $Group.ObjectID
                $GroupMembers += $Object
            }
            return $GroupMembers
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureADContacts)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureADContacts"
        $Data.O365UAzureADContacts = Get-MsolContact -All
    }
    # Below is data that is prepared entirely using data from above (suitable for Word for the most part)
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureTenantDomains)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureTenantDomains (prepared data)"
        $Data.O365AzureTenantDomains = Invoke-Command -ScriptBlock {
            $Domains = @()
            foreach ($Domain in $Data.O365UAzureTenantDomains) {
                $Domains += [PsCustomObject] @{
                    'Domain Name'         = $Domain.Name
                    'Default'             = $Domain.IsDefault
                    'Initial'             = $Domain.IsInitial
                    'Status'              = $Domain.Status
                    'Verification Method' = $Domain.VerificationMethod
                    'Capabilities'        = $Domain.Capabilities
                    'Authentication'      = $Domain.Authentication
                }
            }
            return $Domains
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADGroupMembersUser)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureADGroupMembersUser"
        $Data.O365AzureADGroupMembersUser = Invoke-Command -ScriptBlock {
            $Members = @()
            foreach ($Group in $Data.O365UAzureADGroups) {
                $GroupMembers = $Data.O365UAzureADGroupMembers | Where { $_.GroupObjectId -eq $Group.ObjectId }
                foreach ($GroupMember in $GroupMembers) {
                    $Members += [PsCustomObject] @{
                        "GroupDisplayName"    = $Group.DisplayName
                        "GroupEmail"          = $Group.EmailAddress
                        "GroupEmailSecondary" = $Group.ProxyAddresses -replace 'smtp:', '' -join ','
                        "GroupType"           = $Group.GroupType
                        "MemberDisplayName"   = $GroupMember.DisplayName
                        "MemberEmail"         = $GroupMember.EmailAddress
                        "MemberType"          = $GroupMember.GroupMemberType
                        "LastDirSyncTime"     = $Group.LastDirSyncTime
                        "ManagedBy"           = $Group.ManagedBy
                        "ProxyAddresses"      = $Group.ProxyAddresses
                    }
                }
            }
            return $Members
        }
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADUsersMFA)) {

        $Data.O365AzureADUsersMFA = Invoke-Command -ScriptBlock {
            $AzureUsers = @()
            foreach ($User in $Data.O365UAzureADUsers) {
                $MFAOptions = @{}
                $MFAOptions.AuthAvailable = @()
                foreach ($Auth in $User.StrongAuthenticationMethods) {
                    if ($Auth.IsDefault) {
                        $MFAOptions.AuthDefault = $Auth.MethodType
                    } else {
                        $MFAOptions.AuthAvailable += $Auth.MethodType
                    }
                }

                $AzureUsers += [pscustomobject] @{
                    'UserPrincipalName'             = $User.UserPrincipalName
                    'Display Name'                  = $User.DisplayName

                    'Method Default'                = $MFAOptions.AuthDefault
                    'Method Alternative'            = ($MFAOptions.AuthAvailable | Sort-Object) -join ','


                    'App Authentication Type'       = $User.StrongAuthenticationPhoneAppDetails.AuthenticationType
                    'App Device Name'               = $User.StrongAuthenticationPhoneAppDetails.DeviceName
                    'App Device Tag'                = $User.StrongAuthenticationPhoneAppDetails.DeviceTag
                    'App Device Token'              = $User.StrongAuthenticationPhoneAppDetails.DeviceToken
                    'App Notification Type'         = $User.StrongAuthenticationPhoneAppDetails.NotificationType
                    'App Oath Secret Key'           = $User.StrongAuthenticationPhoneAppDetails.OathSecretKey
                    'App Oath Token Time Drift'     = $User.StrongAuthenticationPhoneAppDetails.OathTokenTimeDrift
                    'App Version'                   = $User.StrongAuthenticationPhoneAppDetails.PhoneAppVersion

                    'User Details Email'            = $User.StrongAuthenticationUserDetails.Email
                    'User Details Phone'            = $User.StrongAuthenticationUserDetails.PhoneNumber
                    'User Details Phone Alt'        = $User.StrongAuthenticationUserDetails.AlternativePhoneNumber
                    'User Details Pin'              = $User.StrongAuthenticationUserDetails.Pin
                    'User Details OldPin'           = $User.StrongAuthenticationUserDetails.OldPin
                    'Strong Password Required'      = $User.StrongPasswordRequired

                    'Requirement Relying Party'     = $User.StrongAuthenticationRequirements.RelyingParty
                    'Requirement Not Issued Before' = $User.StrongAuthenticationRequirements.RememberDevicesNotIssuedBefore
                    'Requirement State'             = $User.StrongAuthenticationRequirements.State

                    # Below needs checking...
                    StrongAuthenticationProofupTime = $User.StrongAuthenticationProofupTime

                }

            }
            return $AzureUsers | Sort-Object 'UserPrincipalName'
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADUsersStatisticsByCountry)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADUsersStatisticsByCountry"
        $Data.O365AzureADUsersStatisticsByCountry = $Data.O365UAzureADUsers | Group-Object Country | Select-Object @{ L = 'Country'; Expression = { if ($_.Name -ne '') { $_.Name } else { 'Unknown' } }} , @{ L = 'Users Count'; Expression = { $_.Count }} | Sort-Object 'Country'
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADUsersStatisticsByCity)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADUsersStatisticsByCity"
        $Data.O365AzureADUsersStatisticsByCity = $Data.O365UAzureADUsers | Group-Object City | Select-Object @{ L = 'City'; Expression = { if ($_.Name -ne '') { $_.Name } else { 'Unknown' } }} , @{ L = 'Users Count'; Expression = { $_.Count }} | Sort-Object 'City'
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADUsersStatisticsByCountryCity)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADUsersStatisticsByCountryCity"
        $Data.O365AzureADUsersStatisticsByCountryCity = $Data.O365UAzureADUsers |  Group-Object Country, City | Select-Object @{ L = 'Country, City'; Expression = { if ($_.Name -ne '') { $_.Name } else { 'Unknown' } }} , @{ L = 'Users Count'; Expression = { $_.Count }} | Sort-Object 'Country, City'
    }
    return $Data
}