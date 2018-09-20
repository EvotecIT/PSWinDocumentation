Add-Type -TypeDefinition @"
    public enum Exchange {
        ExchangeServers,
        ExchangeDatabases,
        ExchangeDatabasesBackup,
        ExchangePublicFolderDatabase,
        ExchangeMailboxes,
        ExchangeMailboxesStatistics
    }
"@