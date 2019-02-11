function Get-ReportO365DistributionGroups {
    [CmdletBinding()]    
    param(
        [switch] $All
    )      
    $O365 = @{}
    $O365.Groups = [System.Collections.Generic.List[PSCustomObject]]::new()
    $O365.GroupsWithProperties = Get-DistributionGroup -ResultSize Unlimited
    $O365.GroupsAllMembers = foreach ($O365Group in $O365.GroupsWithProperties) {
        # This creates new, cleaner groups list
        $O365.Groups.Add( 
            [PSCustomObject] @{ 
                "Group Name"                       = $O365Group.DisplayName 
                "Group Owners"                     = $O365Group.ManagedBy -join ', '
                "Group Primary Email"              = $O365Group.PrimarySmtpAddress
                "Group Emails"                     = Convert-ExchangeEmail -Emails $O365Group.EmailAddresses -AddSeparator -RemoveDuplicates -RemovePrefix

                IsDirSynced                        = $O365Group.IsDirSynced
                MemberJoinRestriction              = $O365Group.MemberJoinRestriction
                MemberDepartRestriction            = $O365Group.MemberDepartRestriction

                GrantSendOnBehalfTo                = $O365Group.GrantSendOnBehalfTo
                MailTip                            = $O365Group.MailTip

                Identity                           = $O365Group.Identity
                SamAccountName                     = $O365Group.SamAccountName
                GroupType                          = $O365Group.GroupType
                WhenCreated                        = $O365Group.WhenCreated
                WhenChanged                        = $O365Group.WhenChanged
                
                Alias                              = $O365Group.Alias
                ModeratedBy                        = $O365Group.ModeratedBy
                ModerationEnabled                  = $O365Group.ModerationEnabled
                HiddenGroupMembershipEnabled       = $O365Group.HiddenGroupMembershipEnabled
             
                              
                HiddenFromAddressListsEnabled      = $O365Group.HiddenFromAddressListsEnabled
                RequireSenderAuthenticationEnabled = $O365Group.RequireSenderAuthenticationEnabled
                RecipientTypeDetails               = $O365Group.RecipientTypeDetails
               
            }   
        )
        # This returns members of groups
        $O365GroupPeople = Get-DistributionGroupMember -Identity $O365Group.GUID.GUID
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