function Get-ReportO365Mailboxes {
    [CmdletBinding()]
    param(
        [string] $Prefix,
        [validateset("Bytes", "KB", "MB", "GB", "TB")][string]$SizeIn = 'MB',
        [alias('Precision')][int]$SizePrecision = 2,
        [alias('ReturnAll')][switch] $All,
        [switch] $SkipAvailability,
        [switch] $GatherPermissions
    )
    $PropertiesMailbox = 'DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'EmailAddresses', 'HiddenFromAddressListsEnabled', 'Identity', 'ExchangeGuid', 'ArchiveGuid', 'ArchiveQuota', 'ArchiveStatus', 'WhenCreated', 'WhenChanged', 'Guid', 'MailboxGUID', 'RecipientTypeDetails'
    #$PropertiesAzure = 'FirstName', 'LastName', 'Country', 'City', 'Department', 'Office', 'UsageLocation', 'Licenses', 'WhenCreated', 'UserPrincipalName', 'ObjectID'
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
    $Object.Mailbox = & "Get-$($Prefix)Mailbox" -ResultSize Unlimited | Select-Object $PropertiesMailbox
    Write-Verbose "Get-ReportO365Mailboxes - Getting all Azure AD users"
    $Object.Azure = Get-MsolUser -All #| Select-Object $PropertiesAzure
    $Object.MailboxStatistics = [System.Collections.Generic.List[object]]::new()
    $Object.MailboxStatisticsArchive = [System.Collections.Generic.List[object]]::new()
    $Object.MailboxPermissions = [System.Collections.Generic.List[PSCustomObject]]::new()
    $Object.MailboxPermissionsAll = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($Mailbox in $Object.Mailbox) {
        Write-Verbose "Get-ReportO365Mailboxes - Processing Mailbox Statistics for Mailbox $($Mailbox.UserPrincipalName)"
        ($Object.MailboxStatistics).Add( (& "Get-$($Prefix)MailboxStatistics" -Identity $Mailbox.Guid.Guid | Select-Object $PropertiesMailboxStats))
        if ($Mailbox.ArchiveStatus -eq "Active") {
            ($Object.MailboxStatisticsArchive).Add((& "Get-$($Prefix)MailboxStatistics" -Identity $Mailbox.Guid.Guid -Archive | Select-Object $PropertiesMailboxStatsArchive))
        }
    }
    
    Write-Verbose "Get-ReportO365Mailboxes - Preparing output data"
    $Object.Output = foreach ($Mailbox in $Object.Mailbox) {
        $Azure = $Object.Azure | Where-Object { $_.UserPrincipalName -eq $Mailbox.UserPrincipalName }
        $MailboxStats = $Object.MailboxStatistics | Where-Object { $_.MailboxGuid.Guid -eq $Mailbox.ExchangeGuid.Guid }
        $MailboxStatsArchive = $Object.MailboxStatisticsArchive | Where-Object { $_.MailboxGuid.Guid -eq $Mailbox.ArchiveGuid.Guid }

        [PSCustomObject][ordered] @{
            DisplayName               = $Mailbox.DisplayName
            UserPrincipalName         = $Mailbox.UserPrincipalName
            FirstName                 = $Azure.FirstName
            LastName                  = $Azure.LastName
            Country                   = $Azure.Country
            City                      = $Azure.City
            Department                = $Azure.Department
            Office                    = $Azure.Office
            UsageLocation             = $Azure.UsageLocation
            License                   = Convert-Office365License -License $Azure.Licenses.AccountSkuID
            UserCreated               = $Azure.WhenCreated

            Blocked                   = $Azure.BlockCredential
            LastSynchronized          = $azure.LastDirSyncTime
            LastPasswordChange        = $Azure.LastPasswordChangeTimestamp
            PasswordNeverExpires      = $Azure.PasswordNeverExpires

            RecipientType             = $Mailbox.RecipientTypeDetails

            PrimaryEmailAddress       = $Mailbox.PrimarySmtpAddress
            AllEmailAddresses         = Convert-ExchangeEmail -Emails $Mailbox.EmailAddresses -Separator ', ' -RemoveDuplicates -RemovePrefix -AddSeparator

            MailboxLogOn              = $MailboxStats.LastLogonTime
            MailboxLogOff             = $MailboxStats.LastLogoffTime

            MailboxSize               = Convert-ExchangeSize -Size $MailboxStats.TotalItemSize -To $SizeIn -Default '' -Precision $SizePrecision

            MailboxItemCount          = $MailboxStats.ItemCount

            MailboxDeletedSize        = Convert-ExchangeSize -Size $MailboxStats.TotalDeletedItemSize -To $SizeIn -Default '' -Precision $SizePrecision
            MailboxDeletedItemsCount  = $MailboxStats.DeletedItemCount

            MailboxHidden             = $Mailbox.HiddenFromAddressListsEnabled
            MailboxCreated            = $Mailbox.WhenCreated # WhenCreatedUTC
            MailboxChanged            = $Mailbox.WhenChanged # WhenChangedUTC

            ArchiveStatus             = $Mailbox.ArchiveStatus
            ArchiveQuota              = Convert-ExchangeSize -Size $Mailbox.ArchiveQuota -To $SizeIn -Default '' -Display
            ArchiveSize               = Convert-ExchangeSize -Size $MailboxStatsArchive.TotalItemSize -To $SizeIn -Default '' -Precision $SizePrecision
            ArchiveItemCount          = Convert-ExchangeItems -Count $MailboxStatsArchive.ItemCount -Default ''

            ArchiveDeletedSize        = Convert-ExchangeSize -Size $MailboxStatsArchive.TotalDeletedItemSize -To $SizeIn -Default '' -Precision $SizePrecision
            ArchiveDeletedItemsCount  = Convert-ExchangeItems -Count $MailboxStatsArchive.DeletedItemCount -Default ''
            # Adding GUID so it's possible to match other data
            OverallProvisioningStatus = $Azure.OverallProvisioningStatus
            ImmutableID               = $Azure.ImmutableID
            Guid                      = $Mailbox.Guid.Guid
            ObjectID                  = $Mailbox.ExternalDirectoryObjectId
        }

        if ($GatherPermissions) {
            $MailboxPermissions = Get-MailboxPermission -Identity $Mailbox.PrimarySmtpAddress.ToString()
            #No non-default permissions found, continue to next mailbox
            if (-not $MailboxPermissions) { continue }
        
            $Permissions = foreach ($Permission in ($MailboxPermissions | Where-Object {($_.User -ne "NT AUTHORITY\SELF") -and ($_.IsInherited -ne $true)}) ) {
                [PSCustomObject] @{
                    DiplayName           = $Mailbox.DisplayName
                    UserPrincipalName    = $Mailbox.UserPrincipalName
                    FirstName            = $Azure.FirstName
                    LastName             = $Azure.LastName
                    RecipientType        = $Mailbox.RecipientTypeDetails
                    PrimaryEmailAddress  = $Mailbox.PrimarySmtpAddress

                    "User With Access"   = $Permission.User
                    "User Access Rights" = ($Permission.AccessRights -join ",")
                }
            }
            if ($null -ne $Permissions) {
                $Object.MailboxPermissions.Add($Permissions)
            }
            $PermissionsAll = foreach ($Permission in $MailboxPermissions) {
                [PSCustomObject] @{
                    DiplayName           = $Mailbox.DisplayName
                    UserPrincipalName    = $Mailbox.UserPrincipalName
                    FirstName            = $Azure.FirstName
                    LastName             = $Azure.LastName
                    RecipientType        = $Mailbox.RecipientTypeDetails
                    PrimaryEmailAddress  = $Mailbox.PrimarySmtpAddress

                    "User With Access"   = $Permission.User
                    "User Access Rights" = ($Permission.AccessRights -join ",")
                    "Inherited"          = $Permission.IsInherited
                    "Deny"               = $Permission.Deny
                    "InheritanceType"    = $Permission.InheritanceType
                }
            }
            if ($null -ne $PermissionsAll) {
                $Object.MailboxPermissionsAll.Add($PermissionsAll)
            }
        }
    }
    if ($All) {
        return $Object
    } else {
        return $Object.Output
    }
}