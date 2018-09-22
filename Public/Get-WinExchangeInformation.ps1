function Get-WinExchangeInformation {
    [CmdletBinding()]
    param(
        $TypesRequired
    )
    $Data = @{}
    $Data.ExchangeServers = Get-ExchangeServer
    $Data.ExchangeDatabases = Invoke-Command -ScriptBlock {
        # Get Exchange Databases
        $Command = @(Get-Command Get-MailboxDatabase)[0]
        if ($Command.Parameters.ContainsKey("IncludePreExchange2010")) {
            $Databases = Get-MailboxDatabase -Status -IncludePreExchange2010
        } elseif ($Command.Parameters.ContainsKey("IncludePreExchange2013")) {
            $Databases = Get-MailboxDatabase -Status -IncludePreExchange2013
        } else {
            $Databases = Get-MailboxDatabase -Status
        }
        return $Databases
    }
    $Data.ExchangeDatabasesBackup = Invoke-Command -ScriptBlock {
        #$Backups = $Data.ExchangeDatabase | Select Name, LastFullBackup, LastIncrementalBackup, LastDifferentialBackup
        $Backups = @()
        $Backups += [ordered] @{
            Name                   = $Data.ExchangeDatabases.Name
            Mounted                = $Data.ExchangeDatabases.Mounted
            LastFullBackup         = if ($Data.ExchangeDatabases.LastFullbackup) { $Data.ExchangeDatabases.LastFullbackup.ToUniversalTime() } else { 'N/A' }
            LastIncrementalBackup  = if ($Data.ExchangeDatabases.LastIncrementalBackup) { $Data.ExchangeDatabases.LastIncrementalBackup.ToUniversalTime() } else { 'N/A' }
            LastDifferentialBackup = if ($Data.ExchangeDatabases.LastDifferentialBackup) {  $Data.ExchangeDatabases.LastDifferentialBackup.ToUniversalTime() } else { 'N/A' }
        }
        return Format-TransposeTable -Object $Backups
    }
    $Data.ExchangePublicFolderDatabase = Invoke-Command -ScriptBlock {
        #Get Public Folder Databases
        $Command = @(Get-Command Get-PublicFolderDatabase)[0]
        if ($Command.Parameters.ContainsKey("IncludePreExchange2010")) {
            $Databases = @(Get-PublicFolderDatabase -Status -IncludePreExchange2010)
        } else {
            $Databases = @(Get-PublicFolderDatabase -Status)
        }
        return $Databases
    }
    $Data.ExchangeMailboxes = Invoke-Command -ScriptBlock {
        $Mailboxes = Get-Mailbox -ResultSize Unlimited
        return $Mailboxes
    }


    $Data.ExchangeMailboxesStatistics = Invoke-Command -ScriptBlock {
        $i = 0
        $ExchangeMailboxesStatistics = @()
        foreach ($Mailbox in $Data.ExchangeMailboxes) {
            $i = $i + 1
            $PercentComplete = $i / $Data.ExchangeMailboxes.Count * 100
            Write-Verbose "Collecting mailbox details Processing mailbox $i of $($Data.ExchangeMailboxes.Count) - $Mailbox Percent Complete $PercentComplete"
            $ExchangeStatistics = $Mailbox | Get-MailboxStatistics | Select-Object TotalItemSize, TotalDeletedItemSize, ItemCount, LastLogonTime, LastLoggedOnUserAccount

            if ($Mailbox.ArchiveDatabase) {
                $ExchangeStatisticsArchive = $Mailbox | Get-MailboxStatistics -Archive | Select-Object TotalItemSize, TotalDeletedItemSize, ItemCount
            } else {
                $ExchangeStatisticsArchive = "n/a"
            }

            $ExchangeStatisticsInbox = Get-MailboxFolderStatistics $Mailbox.Identity -FolderScope Inbox | Where-Object {$_.FolderPath -eq "/Inbox"}
            $ExchangeStatisticsSent = Get-MailboxFolderStatistics $Mailbox.Identity -FolderScope SentItems | Where-Object {$_.FolderPath -eq "/Sent Items"}
            $ExchangeStatisticsDeleted = Get-MailboxFolderStatistics $Mailbox.Identity -FolderScope DeletedItems | Where-Object {$_.FolderPath -eq "/Deleted Items"}

            $ExchangeUser = Get-User $Mailbox.Identity
            $ActiveDirectoryUser = Get-ADUser $Mailbox.SamAccountName -Properties Enabled, AccountExpirationDate

            $ExchangeDatabasePrimary = $Data.ExchangeMailboxes | Where-Object {$_.Name -eq $Mailbox.Database.Name}
            $ExchangeDatabaseArchive = $Data.ExchangeMailboxes | Where-Object {$_.Name -eq $Mailbox.ArchiveDatabase.Name}

            $UserObject = [PSCustomObject]@{
                "DisplayName"                   = $Mailbox.DisplayName
                "Mailbox Type"                  = $Mailbox.RecipientTypeDetails
                "Title"                         = $ExchangeUser.Title
                "Department"                    = $ExchangeUser.Department
                "Office"                        = $ExchangeUser.Office

                "Total Mailbox Size"            = '' # (($stats.TotalItemSize.Value + $stats.TotalDeletedItemSize.Value))
                "Mailbox Size"                  = $ExchangeStatistics.TotalItemSize.Value
                "Mailbox Recoverable Item Size" = $ExchangeStatistics.TotalDeletedItemSize.Value
                "Mailbox Items"                 = $ExchangeStatistics.ItemCount
                "Inbox Folder Size"             = $ExchangeStatisticsInbox.FolderandSubFolderSize
                "Sent Items Folder Size"        = $ExchangeStatisticsSent.FolderandSubFolderSize
                "Deleted Items Folder Size"     = $ExchangeStatisticsDeleted.FolderandSubFolderSize

                "Total Archive Size"            = if ($ExchangeStatisticsArchive -eq "n/a") { 'n/a' } else { ($ExchangeStatisticsArchive.TotalItemSize.Value + $ExchangeStatisticsArchive.TotalDeletedItemSize.Value) }
                "Archive Size"                  = if ($ExchangeStatisticsArchive -eq "n/a") { 'n/a' } else { $ExchangeStatisticsArchive.TotalItemSize.Value }
                "Archive Deleted Item Size"     = if ($ExchangeStatisticsArchive -eq "n/a") { 'n/a' } else { $ExchangeStatisticsArchive.TotalDeletedItemSize.Value }
                "Archive Items"                 = if ($ExchangeStatisticsArchive -eq "n/a") { 'n/a' } else { $ExchangeStatisticsArchive.ItemCount }

                "Audit Enabled"                 = $Mailbox.AuditEnabled
                "Email Address Policy Enabled"  = $Mailbox.EmailAddressPolicyEnabled
                "Hidden From Address Lists"     = $Mailbox.HiddenFromAddressListsEnabled
                "Use Database Quota Defaults"   = $Mailbox.UseDatabaseQuotaDefaults

                "Issue Warning Quota"           = if ($Mailbox.UseDatabaseQuotaDefaults) { $ExchangeDatabasePrimary.IssueWarningQuota } else { $Mailbox.IssueWarningQuota }
                "Prohibit Send Quota"           = if ($Mailbox.UseDatabaseQuotaDefaults) { $ExchangeDatabasePrimary.ProhibitSendQuota } else { $Mailbox.ProhibitSendQuota }
                "Prohibit Send Receive Quota"   = if ($Mailbox.UseDatabaseQuotaDefaults) { $ExchangeDatabasePrimary.ProhibitSendReceiveQuota } else { $Mailbox.ProhibitSendReceiveQuota }


                "Account Enabled"               = $ActiveDirectoryUser.Enabled
                "Account Expires"               = $ActiveDirectoryUser.AccountExpirationDate
                "Last Mailbox Logon"            = $ExchangeStatistics.LastLogonTime
                "Last Logon By"                 = $ExchangeStatistics.LastLoggedOnUserAccount


                "Primary Mailbox Database"      = $Mailbox.Database
                "Primary Server/DAG"            = $ExchangeDatabasePrimary.MasterServerOrAvailabilityGroup

                "Archive Mailbox Database"      = $Mailbox.ArchiveDatabase
                "Archive Server/DAG"            = $ExchangeDatabaseArchive.MasterServerOrAvailabilityGroup

                "Primary Email Address"         = $Mailbox.PrimarySMTPAddress
                "Organizational Unit"           = $ExchangeUser.OrganizationalUnit
            }
            $ExchangeMailboxesStatistics += $UserObject
        }
        return $ExchangeMailboxesStatistics
    }
    return $Data
}