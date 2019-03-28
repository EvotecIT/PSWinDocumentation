function Get-ReportO365Licenses {
    [CmdletBinding()]
    param(
        [switch] $All
    )

    $Mailboxes = get-mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Select-Object UserPrincipalName, PrimarySMTPAddress, SamAccountName, DisplayName, Name, Identity

    $i = 1
    $InboxRules = @(
        foreach ($Mailbox in $Mailboxes) {
            Write-Color -Text "$($i) of $($Mailboxes.Count)", ' | ', "$($Mailbox.UserPrincipalName)" -Color Yellow, White, Blue, White
            Get-InboxRule -Mailbox $mailbox.UserPrincipalName | Select-Object *
            $i++
        }
    )
    $InboxRulesForwarding = @(
        foreach ($Mailbox in $Mailboxes) {
            $UserRules = $InboxRules | Where-Object { ($Mailbox.Identity -eq $_.MailboxOwnerID) -and (($null -ne $_.ForwardTo) -or ($null -ne $_.ForwardAsAttachmentTo) -or ($null -ne $_.RedirectsTo)) }
            foreach ($Rule in $UserRules) {
                [pscustomobject][ordered] @{
                    UserPrincipalName     = $Mailbox.UserPrincipalName
                    DisplayName           = $Mailbox.DisplayName
                    RuleName              = $Rule.Name
                    Description           = $Rule.Description
                    Enabled               = $Rule.Enabled
                    Priority              = $Rule.Priority
                    ForwardTo             = $Rule.ForwardTo
                    ForwardAsAttachmentTo = $Rule.ForwardAsAttachmentTo
                    RedirectTo            = $Rule.RedirectTo
                    DeleteMessage         = $Rule.DeleteMessage
                }
            }
        }
    )
    #$Mailboxes | ConvertTo-Excel -FilePath $Configuration.OutputFile -ExcelWorkSheetName 'All Mailboxes' -AutoFilter -AutoFit
    #$InboxRules | Select-Object * -ExcludeProperty PSComputerName, RunspaceID, PSShowComputerName, PSComputerName, IsValid, ObjectState | ConvertTo-Excel -FilePath $Configuration.OutputFile -ExcelWorkSheetName 'Inbox Rules' -AutoFilter -AutoFit
    #$InboxRulesForwarding | ConvertTo-Excel -FilePath $Configuration.OutputFile -ExcelWorkSheetName 'Inbox Rules with Forwarding' -AutoFilter -AutoFit
    if ($All) {
        $InboxRules | Select-Object * -ExcludeProperty PSComputerName, RunspaceID, PSShowComputerName, PSComputerName, IsValid, ObjectState
    } else {
        $InboxRulesForwarding
    }
}


#$InboxRules = Get-ReportO365Licenses
#$InboxRules | ConvertTo-Excel -FilePath $Env:USERPROFILE\Desktop\InboxRules.xlsx -AutoFilter -AutoFit -FreezeTopRowFirstColumn

#$InboxRules |  FL *