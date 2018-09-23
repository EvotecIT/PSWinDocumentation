function Get-WinO365Azure {
    [CmdletBinding()]
    param(
        $TypesRequired
    )
    $Data = [ordered] @{}
    if ($TypesRequired -eq $null) {
        Write-Verbose 'Get-WinO365Azure - TypesRequired is null. Getting all Exchange types.'
        $TypesRequired = Get-Types -Types ([O365])  # Gets all types
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureLicensing)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureLicensing"
        $Data.O365AzureLicensing = Get-MsolAccountSku
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureTenantDomains)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureTenantDomains"
        $Data.O365AzureTenantDomains = Get-MsolDomain
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureSubscription)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureSubscription"
        $Data.O365AzureSubscription = Get-MsolSubscription
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADUsers)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADUsers"
        $Data.O365AzureADUsers = Get-MsolUser -All
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADUsersDeleted)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADUsersDeleted"
        $Data.O365AzureADUsersDeleted = Get-MsolUser -ReturnDeletedUsers
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADGroups)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADGroups"
        $Data.O365AzureADGroups = Get-MsolGroup -All
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADGroupMembers)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADGroupMembers"
        $Data.O365AzureADGroupMembers = Invoke-Command -ScriptBlock {
            $GroupMembers = @()
            foreach ($Group in $Data.Groups) {
                $Object = Get-MsolGroupMember -GroupObjectId $Group.ObjectId -All
                $Object | Add-Member -MemberType NoteProperty -Name "GroupObjectID" -Value $Group.ObjectID
                $GroupMembers += $Object
            }
            return $GroupMembers
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADContacts)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADContacts"
        $Data.O365AzureADContacts = Get-MsolContact -All
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365AzureADGroupMembersUser)) {
        Write-Verbose "Get-WinO365Azure - Getting O365AzureADGroupMembersUser"
        $Data.O365AzureADGroupMembersUser = Invoke-Command -ScriptBlock {
            $Members = @()
            foreach ($Group in $Data.O365AzureADGroups) {
                $GroupMembers = $Data.O365AzureADGroupMembers | Where { $_.GroupObjectId -eq $Group.ObjectId }
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