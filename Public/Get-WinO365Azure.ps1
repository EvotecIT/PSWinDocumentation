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

                $Licenses += [PSCustomObject] @{
                    Name                 = $($Global:O365SKU).Item("$($License.SkuPartNumber)")
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
            return $Licenses
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureTenantDomains, [O365]::O365AzureTenantDomains)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureTenantDomains"
        $Data.O365UAzureTenantDomains = Get-MsolDomain
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureSubscription)) {
        Write-Verbose "Get-WinO365Azure - Getting O365UAzureSubscription"
        $Data.O365UAzureSubscription = Get-MsolSubscription
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UAzureADUsers)) {
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
    return $Data
}