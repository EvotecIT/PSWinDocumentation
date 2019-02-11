function Get-ReportO365UnifiedGroups {
    [CmdletBinding()]    
    param(
        [switch] $All
    )
    $O365 = @{}
    $O365.Groups = [System.Collections.Generic.List[PSCustomObject]]::new()
    $O365.GroupsWithProperties = Get-UnifiedGroup -ResultSize Unlimited -IncludeAllProperties
    $O365.GroupsAllMembers = foreach ($O365Group in $O365.GroupsWithProperties) {
        # This creates new, cleaner groups list
        $O365.Groups.Add( 
            [PSCustomObject] @{ 
                "Group Name"                           = $O365Group.DisplayName 
                "Group Owners"                         = $O365Group.ManagedBy -join ', '
                "Group Primary Email"                  = $O365Group.PrimarySmtpAddress
                "Group Emails"                         = Convert-ExchangeEmail -Emails $O365Group.EmailAddresses -AddSeparator -RemoveDuplicates -RemovePrefix
                Identity                               = $O365Group.Identity
                WhenCreated                            = $O365Group.WhenCreated
                WhenChanged                            = $O365Group.WhenChanged
                
                Alias                                  = $O365Group.Alias
                ModerationEnabled                      = $O365Group.ModerationEnabled
                AccessType                             = $O365Group.AccessType
                AutoSubscribeNewMembers                = $O365Group.AutoSubscribeNewMembers
                AlwaysSubscribeMembersToCalendarEvents = $O365Group.AlwaysSubscribeMembersToCalendarEvents
                CalendarMemberReadOnly                 = $O365Group.CalendarMemberReadOnly
                HiddenGroupMembershipEnabled           = $O365Group.HiddenGroupMembershipEnabled
                SubscriptionEnabled                    = $O365Group.SubscriptionEnabled
                              
                HiddenFromExchangeClientsEnabled       = $O365Group.HiddenFromExchangeClientsEnabled
                InboxUrl                               = $O365Group.InboxUrl
                SharePointSiteUrl                      = $O365Group.SharePointSiteUrl
                SharePointDocumentsUrl                 = $O365Group.SharePointDocumentsUrl
                SharePointNotebookUrl                  = $O365Group.SharePointNotebookUrl
               
            }   
        )
        # This returns members of groups
        $O365GroupPeople = Get-UnifiedGroupLinks -Identity $O365Group.Guid.Guid -LinkType Members 
        foreach ($O365Member in $O365GroupPeople) { 
            [PSCustomObject] @{ 
                "Group Name"          = $O365Group.DisplayName 
                "Group Primary Email" = $O365Group.PrimarySmtpAddress
                "Group Emails"        = Convert-ExchangeEmail -Emails $O365Group.EmailAddresses -AddSeparator -RemoveDuplicates -RemovePrefix
                "Group Owners"        = $O365Group.ManagedBy -join ', '
                "Member Name"         = $O365Member.Name 
                "Member E-Mail"       = $O365Member.PrimarySMTPAddress 
                "Recipient Type"      = $O365Member.RecipientType 
            }                   
        } 
    }              
    if ($All) {
        $O365
    } else {
        $O365.GroupsAllMembers
    }
} 