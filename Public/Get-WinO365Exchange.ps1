function Get-WinO365Exchange {
    [CmdletBinding()]
    param(
        $TypesRequired
    )
    $Data = [ordered] @{}
    if ($TypesRequired -eq $null) {
        Write-Verbose 'Get-WinO365Exchange - TypesRequired is null. Getting all O365UExchange types.'
        $TypesRequired = Get-Types -Types ([O365])  # Gets all types
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @(
            [O365]::O365UExchangeMailBoxes,
            [O365]::O365UExchangeMailboxesJunk,
            [O365]::O365UExchangeMailboxesRooms,
            [O365]::O365UExchangeMailboxesEquipment,
            [O365]::O365UExchangeInboxRules
        )) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeMailBoxes"
        $Data.O365UExchangeMailBoxes = Get-O365Mailbox -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeMailUsers)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeMailUsers"
        $Data.O365UExchangeMailUsers = Get-O365MailUser -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeUsers)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeUsers"
        $Data.O365UExchangeUsers = Get-O365User -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeRecipients)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeRecipients"
        $Data.O365UExchangeRecipients = Get-O365Recipient -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeRecipientsPermissions)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeRecipientsPermissions"
        $Data.O365UExchangeRecipientsPermissions = Get-O365RecipientPermission -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeGroupsDistribution, [O365]::O365UExchangeGroupsDistributionMembers)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeGroupsDistribution"
        $Data.O365UExchangeGroupsDistribution = Get-O365DistributionGroup -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeGroupsDistributionDynamic)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeGroupsDistributionDynamic"
        $Data.O365UExchangeGroupsDistributionDynamic = Get-O365DynamicDistributionGroup -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeGroupsDistributionMembers)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeGroupsDistributionMembers"
        $Data.O365UExchangeGroupsDistributionMembers = Invoke-Command -ScriptBlock {
            $GroupMembers = @()
            foreach ($Group in $Data.O365UExchangeGroupsDistribution) {
                $Object = Get-O365DistributionGroupMember -Identity $Group.PrimarySmtpAddress -ResultSize unlimited
                $Object | Add-Member -MemberType NoteProperty -Name "GroupGUID" -Value $Group.GUID
                $Object | Add-Member -MemberType NoteProperty -Name "GroupPrimarySmtpAddress" -Value $Group.PrimarySmtpAddress
                $Object | Add-Member -MemberType NoteProperty -Name "GroupIdentity" -Value $Group.Identity
                $GroupMembers += $Object
            }
            return $GroupMembers
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeMailboxesJunk)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeMailboxesJunk"
        $Data.O365UExchangeMailboxesJunk = Invoke-Command -ScriptBlock {
            $Output = @()
            foreach ($Mailbox in $Data.O365UExchangeMailBoxes) {
                if ($Mailbox.PrimarySmtpAddress -ne $null) {
                    #Write-Verbose "O365UExchangeMailboxesJunk - $($Mailbox.PrimarySmtpAddress)"
                    $Object = Get-O365MailboxJunkEmailConfiguration -Identity $Mailbox.PrimarySmtpAddress -ResultSize unlimited
                    if ($Object) {
                        $Object | Add-Member -MemberType NoteProperty -Name "MailboxPrimarySmtpAddress" -Value $Mailbox.PrimarySmtpAddress
                        $Object | Add-Member -MemberType NoteProperty -Name "MailboxAlias" -Value $Mailbox.Alias
                        $Object | Add-Member -MemberType NoteProperty -Name "MailboxGUID" -Value $Mailbox.GUID
                        $Output += $Object
                    }
                }
            }
            return $Output
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeContacts)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeContacts"
        $Data.O365UExchangeContacts = Get-O365Contact -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeInboxRules)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeContacts"
        $Data.O365UExchangeInboxRules = Invoke-Command -ScriptBlock {
            $InboxRules = @()
            foreach ($Mailbox in $Data.O365UExchangeMailBoxes) {
                $InboxRules += Get-O365InboxRule -Mailbox $Mailbox.UserPrincipalName
            }
            return $InboxRules
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeContacts)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeContacts"
        $Data.O365ExchangeInboxRules = Invoke-Command -ScriptBlock {
            return $Data.O365UExchangeInboxRules
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeContacts)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeContacts"
        $Data.O365ExchangeInboxRules = Invoke-Command -ScriptBlock {
            return $Data.O365UExchangeInboxRules
        }
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeContactsMail)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeContactsMail"
        $Data.O365UExchangeContactsMail = Get-O365MailContact -ResultSize unlimited
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeMailboxesRooms, [O365]::O365UExchangeRoomsCalendarProcessing)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeMailboxesRooms"
        $Data.O365UExchangeMailboxesRooms = $Data.O365UExchangeMailBoxes | Where { $_.RecipientTypeDetails -eq 'RoomMailbox' }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeMailboxesEquipment, [O365]::O365UExchangeEquipmentCalendarProcessing)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeMailboxesEquipment"
        $Data.O365UExchangeMailboxesEquipment = $Data.O365UExchangeMailBoxes | Where { $_.RecipientTypeDetails -eq 'EquipmentMailbox' }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeRoomsCalendarProcessing)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeRoomsCalendarProcessing"
        $Data.O365UExchangeRoomsCalendarProcessing = Invoke-Command -ScriptBlock {
            $Output = @()
            foreach ($Mailbox in $Data.O365UExchangeMailboxesRooms) {
                $Object = Get-O365CalendarProcessing -Identity $Mailbox.PrimarySmtpAddress -ResultSize unlimited
                if ($Object) {
                    $Object | Add-Member -MemberType NoteProperty -Name "MailboxPrimarySmtpAddress" -Value $Mailbox.PrimarySmtpAddress
                    $Object | Add-Member -MemberType NoteProperty -Name "MailboxAlias" -Value $Mailbox.Alias
                    $Object | Add-Member -MemberType NoteProperty -Name "MailboxGUID" -Value $Mailbox.GUID
                    $Output += $Object
                }
            }

        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([O365]::O365UExchangeEquipmentCalendarProcessing)) {
        Write-Verbose "Get-WinO365Exchange - Getting O365UExchangeEquipmentCalendarProcessing"
        $Data.O365UExchangeEquipmentCalendarProcessing = Invoke-Command -ScriptBlock {
            $Output = @()
            foreach ($Mailbox in $Data.O365UExchangeMailboxesEquipment) {
                $Object = Get-O365CalendarProcessing -Identity $Mailbox.PrimarySmtpAddress -ResultSize unlimited
                if ($Object) {
                    $Object | Add-Member -MemberType NoteProperty -Name "MailboxPrimarySmtpAddress" -Value $Mailbox.PrimarySmtpAddress
                    $Object | Add-Member -MemberType NoteProperty -Name "MailboxAlias" -Value $Mailbox.Alias
                    $Object | Add-Member -MemberType NoteProperty -Name "MailboxGUID" -Value $Mailbox.GUID
                    $Output += $Object
                }
            }
        }
    }
    return $Data
}
