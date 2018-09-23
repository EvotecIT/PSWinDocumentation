function Get-WinExchangeInformation {
    [CmdletBinding()]
    param(
        $TypesRequired
    )
    $Data = [ordered] @{}
    if ($TypesRequired -eq $null) {
        Write-Verbose 'Get-WinExchangeInformation - TypesRequired is null. Getting all Exchange types.'
        $TypesRequired = Get-Types -Types ([Exchange])  # Gets all types
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([Exchange]::ExchangeUServers)) {
        $Data.ExchangeUServers = Get-ExchangeServer
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([Exchange]::ExchangeUDatabases)) {
        $Data.ExchangeUDatabases = Invoke-Command -ScriptBlock {
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
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([Exchange]::ExchangeUPublicFolderDatabase)) {
        $Data.ExchangeUPublicFolderDatabase = Invoke-Command -ScriptBlock {
            #Get Public Folder Databases
            $Command = @(Get-Command Get-PublicFolderDatabase)[0]
            if ($Command.Parameters.ContainsKey("IncludePreExchange2010")) {
                $Databases = @(Get-PublicFolderDatabase -Status -IncludePreExchange2010)
            } else {
                $Databases = @(Get-PublicFolderDatabase -Status)
            }
            return $Databases
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([Exchange]::ExchangeUMailboxes)) {
        $Data.ExchangeUMailboxes = Invoke-Command -ScriptBlock {
            $Mailboxes = Get-Mailbox -ResultSize Unlimited
            return $Mailboxes
        }
    }


    # Below data is prepared data
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([Exchange]::ExchangeDatabasesBackup)) {
        $Data.ExchangeDatabasesBackup = Invoke-Command -ScriptBlock {

            $Backups = @()
            foreach ($DB in $Data.ExchangeUDatabases) {
                $Backups += [pscustomobject] @{
                    Name                   = $DB.Name
                    Mounted                = $DB.Mounted
                    LastFullBackup         = if ($DB.LastFullbackup) { $DB.LastFullbackup.ToUniversalTime() } else { 'N/A' }
                    LastIncrementalBackup  = if ($DB.LastIncrementalBackup) { $DB.LastIncrementalBackup.ToUniversalTime() } else { 'N/A' }
                    LastDifferentialBackup = if ($DB.LastDifferentialBackup) {  $DB.LastDifferentialBackup.ToUniversalTime() } else { 'N/A' }
                }
            }
            return $Backups
        }
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([Exchange]::ExchangeMailboxesStatistics)) {
        $Data.ExchangeMailboxesStatistics = Invoke-Command -ScriptBlock {
            $i = 0
            $ExchangeMailboxesStatistics = @()
            foreach ($Mailbox in $Data.ExchangeUMailboxes) {
                $i = $i + 1
                $PercentComplete = $i / $Data.ExchangeUMailboxes.Count * 100
                Write-Verbose "Collecting mailbox details Processing mailbox $i of $($Data.ExchangeUMailboxes.Count) - $Mailbox Percent Complete $PercentComplete"
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

                $ExchangeDatabasePrimary = $Data.ExchangeUMailboxes | Where-Object {$_.Name -eq $Mailbox.Database.Name}
                $ExchangeDatabaseArchive = $Data.ExchangeUMailboxes | Where-Object {$_.Name -eq $Mailbox.ArchiveDatabase.Name}

                $UserObject = [PSCustomObject]@{
                    "DisplayName"                   = $Mailbox.DisplayName
                    "Mailbox Type"                  = $Mailbox.RecipientTypeDetails
                    "Title"                         = $ExchangeUser.Title
                    "Department"                    = $ExchangeUser.Department
                    "Office"                        = $ExchangeUser.Office

                    #"Total Mailbox Size"            = (($ExchangeStatistics.TotalItemSize.Value + $ExchangeStatistics.TotalDeletedItemSize.Value))
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
    }
    return $Data
}