Add-Type -TypeDefinition @"
    public enum O365 {
        O365ExchangeContacts,
        O365ExchangeRoomsCalendarPrcessing,
        O365ExchangeMailboxesJunk,
        O365ExchangeContactsMail,
        O365ExchangeGroupsDistribution,
        O365ExchangeEquipmentCalendarProcessing,
        O365ExchangeGroupsDistributionMembers,
        O365ExchangeRecipients,
        O365ExchangeMailboxesRooms,
        O365ExchangeUsers,
        O365ExchangeMailboxesEquipment,
        O365ExchangeGroupsDistributionDynamic,
        O365ExchangeRecipientsPermissions,
        O365ExchangeMailUsers,
        O365ExchangeMailBoxes,
        O365AzureLicensing,
        O365AzureTenantDomains,
        O365AzureSubscription,
        O365AzureADUsers,
        O365AzureADUsersDeleted,
        O365AzureADGroups,
        O365AzureADGroupMembersUser,
        O365AzureADGroupMembers,
        O365AzureADContacts
    }
"@