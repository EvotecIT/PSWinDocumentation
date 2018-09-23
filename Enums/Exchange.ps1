Add-Type -TypeDefinition @"
    public enum Exchange {
        // Clean On-Premises Exchange Data
        ExchangeUServers,
        ExchangeUDatabases,
        ExchangeUPublicFolderDatabase,
        ExchangeUMailboxes,

        // Prepared On-Premises Exchange Data
        ExchangeDatabasesBackup,
        ExchangeMailboxesStatistics
    }
"@