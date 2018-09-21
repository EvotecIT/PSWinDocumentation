function Get-WinO365Azure {
    [CmdletBinding()]
    $Data = [ordered] @{}
    Write-Verbose "Get-WinO365Azure - Getting O365AzureLicensing"
    $Data.O365AzureLicensing = Get-MsolAccountSku
    Write-Verbose "Get-WinO365Azure - Getting O365AzureLicensing"
    $Data.O365AzureTenantDomains = Get-MsolDomain
    Write-Verbose "Get-WinO365Azure - Getting O365AzureLicensing"
    $Data.O365AzureSubscription = Get-MsolSubscription
    Write-Verbose "Get-WinO365Azure - Getting O365AzureADUsers"
    $Data.O365AzureADUsers = Get-MsolUser -All
    Write-Verbose "Get-WinO365Azure - Getting O365AzureADUsersDeleted"
    $Data.O365AzureADUsersDeleted = Get-MsolUser -ReturnDeletedUsers
    Write-Verbose "Get-WinO365Azure - Getting O365AzureADGroups"
    $Data.O365AzureADGroups = Get-MsolGroup -All
    Write-Verbose "Get-WinO365Azure - Getting O365AzureADGroupMembers"
    <#
    $Data.O365AzureADGroupMembers = Invoke-Command -ScriptBlock {
        $GroupMembers = @()
        foreach ($Group in $Data.Groups) {
            $Object = Get-MsolGroupMember -GroupObjectId $Group.ObjectId -All
            $Object | Add-Member -MemberType NoteProperty -Name "GroupObjectID" -Value $Group.ObjectID
            $GroupMembers += $Object
        }
        return $GroupMembers
    }
    #>
    Write-Verbose "Get-WinO365Azure - Getting O365AzureADContacts"
    $Data.O365AzureADContacts = Get-MsolContact -All
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
    return $Data
}