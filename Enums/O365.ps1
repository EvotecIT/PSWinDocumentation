Add-Type -TypeDefinition @"
    public enum O365 {
        // Clean O365 Exchange Data
        O365UExchangeContacts,
        O365UExchangeRoomsCalendarPrcessing,
        O365UExchangeMailboxesJunk,
        O365UExchangeContactsMail,
        O365UExchangeGroupsDistribution,
        O365UExchangeEquipmentCalendarProcessing,
        O365UExchangeGroupsDistributionMembers,
        O365UExchangeRecipients,
        O365UExchangeMailboxesRooms,
        O365UExchangeUsers,
        O365UExchangeMailboxesEquipment,
        O365UExchangeGroupsDistributionDynamic,
        O365UExchangeRecipientsPermissions,
        O365UExchangeMailUsers,
        O365UExchangeMailBoxes,
        O365UExchangeInboxRules,

        // Clean O365 Azure Data
        O365UAzureLicensing,
        O365UAzureTenantDomains,
        O365UAzureSubscription,
        O365UAzureADUsers,
        O365UAzureADUsersDeleted,
        O365UAzureADGroups,
        O365UAzureADContacts,
        O365UAzureADGroupMembers,

        // Prepared O365 Azure Data
        O365AzureLicensing,
        O365AzureSubscription,
        O365AzureTenantDomains,

        O365AzureADGroupMembersUser,

        O365AzureADUsersMFA,
        O365AzureADUsersStatisticsByCountry,
        O365AzureADUsersStatisticsByCity,
        O365AzureADUsersStatisticsByCountryCity
    }
"@