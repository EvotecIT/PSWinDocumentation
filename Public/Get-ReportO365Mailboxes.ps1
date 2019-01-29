function Get-ReportO365Mailboxes {
    [CmdletBinding()]
    param(
        [string] $Prefix,
        [validateset("Bytes", "KB", "MB", "GB", "TB")][string]$SizeIn = 'MB',
        [alias('Precision')][int]$SizePrecision = 2,
        [switch] $ReturnAll,
        [switch] $SkipAvailability
    )
    $PropertiesMailbox = 'DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'EmailAddresses', 'HiddenFromAddressListsEnabled', 'Identity', 'ExchangeGuid', 'ArchiveGuid', 'ArchiveQuota', 'ArchiveStatus', 'WhenCreated', 'WhenChanged', 'Guid', 'MailboxGUID'
    $PropertiesAzure = 'FirstName', 'LastName', 'Country', 'City', 'Department', 'Office', 'UsageLocation', 'Licenses', 'WhenCreated', 'UserPrincipalName', 'ObjectID'
    $PropertiesMailboxStats = 'DisplayName', 'LastLogonTime', 'LastLogoffTime', 'TotalItemSize', 'ItemCount', 'TotalDeletedItemSize', 'DeletedItemCount', 'OwnerADGuid', 'MailboxGuid'
    $PropertiesMailboxStatsArchive = 'DisplayName', 'TotalItemSize', 'ItemCount', 'TotalDeletedItemSize', 'DeletedItemCount', 'OwnerADGuid', 'MailboxGuid'

    if ($SkipAvailability) {
        $Commands = Test-AvailabilityCommands -Commands "Get-$($Prefix)Mailbox", "Get-$($Prefix)MsolUser", "Get-$($Prefix)MailboxStatistics"
        if ($Commands -contains $false) {
            Write-Warning "Get-ReportO365Mailboxes - One of commands Get-$($Prefix)Mailbox, Get-$($Prefix)MsolUser, Get-$($Prefix)MailboxStatistics is not available. Make sure connectivity to Office 365 exists."
            return 
        }
    }

    $Object = [ordered] @{}
    Write-Verbose "Get-ReportO365Mailboxes - Getting all mailboxes"
    $Object.Mailboxes = & "Get-$($Prefix)Mailbox" -ResultSize Unlimited | Select-Object $PropertiesMailbox
    Write-Verbose "Get-ReportO365Mailboxes - Getting all Azure AD users"
    $Object.Azure = Get-MsolUser -All | Select-Object $PropertiesAzure
    $Object.MailboxStatistics = [System.Collections.Generic.List[object]]::new()
    $Object.MailboxStatisticsArchive = [System.Collections.Generic.List[object]]::new()
    foreach ($Mailbox in $Object.Mailboxes) {
        Write-Verbose "Get-ReportO365Mailboxes - Processing Mailbox Statistics for Mailbox $($Mailbox.UserPrincipalName)"
        ($Object.MailboxStatistics).Add( (& "Get-$($Prefix)MailboxStatistics" -Identity $Mailbox.Guid.Guid | Select-Object $PropertiesMailboxStats))
        if ($Mailbox.ArchiveStatus -eq "Active") {
            ($Object.MailboxStatisticsArchive).Add((& "Get-$($Prefix)MailboxStatistics" -Identity $Mailbox.Guid.Guid -Archive | Select-Object $PropertiesMailboxStatsArchive))
        }
    }

    Write-Verbose "Get-ReportO365Mailboxes - Preparing output data"
    $Object.Output = foreach ($Mailbox in $Object.Mailboxes) {
        $Azure = $Object.Azure | Where-Object { $_.UserPrincipalName -eq $Mailbox.UserPrincipalName }
        $MailboxStats = $Object.MailboxStatistics | Where-Object { $_.MailboxGuid.Guid -eq $Mailbox.ExchangeGuid.Guid }
        $MailboxStatsArchive = $Object.MailboxStatisticsArchive | Where-Object { $_.MailboxGuid.Guid -eq $Mailbox.ArchiveGuid.Guid }

        [PSCustomObject][ordered] @{
            DiplayName               = $Mailbox.DisplayName
            UserPrincipalName        = $Mailbox.UserPrincipalName
            FirstName                = $Azure.FirstName
            LastName                 = $Azure.LastName
            Country                  = $Azure.Country
            City                     = $Azure.City
            Department               = $Azure.Department
            Office                   = $Azure.Office
            UsageLocation            = $Azure.UsageLocation
            License                  = Convert-Office365License -License $Azure.Licenses.AccountSkuID
            UserCreated              = $Azure.WhenCreated

            PrimaryEmailAddress      = $Mailbox.PrimarySmtpAddress
            AllEmailAddresses        = Convert-ExchangeEmail -Emails $Mailbox.EmailAddresses -Separator ', ' -RemoveDuplicates -RemovePrefix -AddSeparator

            MailboxLogOn             = $MailboxStats.LastLogonTime
            MailboxLogOff            = $MailboxStats.LastLogoffTime

            MailboxSize              = Convert-ExchangeSize -Size $MailboxStats.TotalItemSize -To $SizeIn -Default '' -Precision $SizePrecision

            MailboxItemCount         = $MailboxStats.ItemCount

            MailboxDeletedSize       = Convert-ExchangeSize -Size $MailboxStats.TotalDeletedItemSize -To $SizeIn -Default '' -Precision $SizePrecision
            MailboxDeletedItemsCount = $MailboxStats.DeletedItemCount

            MailboxHidden            = $Mailbox.HiddenFromAddressListsEnabled
            MailboxCreated           = $Mailbox.WhenCreated # WhenCreatedUTC
            MailboxChanged           = $Mailbox.WhenChanged # WhenChangedUTC

            ArchiveStatus            = $Mailbox.ArchiveStatus
            ArchiveQuota             = Convert-ExchangeSize -Size $Mailbox.ArchiveQuota -To $SizeIn -Default '' -Display
            ArchiveSize              = Convert-ExchangeSize -Size $MailboxStatsArchive.TotalItemSize -To $SizeIn -Default '' -Precision $SizePrecision
            ArchiveItemCount         = Convert-ExchangeItems -Count $MailboxStatsArchive.ItemCount -Default ''

            ArchiveDeletedSize       = Convert-ExchangeSize -Size $MailboxStatsArchive.TotalDeletedItemSize -To $SizeIn -Default '' -Precision $SizePrecision
            ArchiveDeletedItemsCount = Convert-ExchangeItems -Count $MailboxStatsArchive.DeletedItemCount -Default ''
            # Adding GUID so it's possible to match other data
            Guid                     = $Mailbox.Guid.Guid
            ObjectID                 = $Mailbox.ExternalDirectoryObjectId
        }
    }
    if ($ReturnAll) {
        return $Object
    } else {
        return $Object.Output
    }
}