function Show-GroupMember {
    [cmdletBinding()]
    param(
        [string[]] $GroupName,
        [string] $FilePath,
        [ValidateSet('Default', 'Hierarchical', 'Both')][string] $Show = 'Both',
        [ValidateSet('Default', 'Hierarchical', 'Both')][string] $RemoveAppliesTo = 'Both',
        [switch] $RemoveComputers,
        [switch] $RemoveUsers,
        [switch] $RemoveOther
    )
    New-HTML -TitleText "Group Membership for $GroupName" {
        New-HTMLSectionStyle -BorderRadius 0px -HeaderBackGroundColor Grey -RemoveShadow
        New-HTMLTableOption -DataStore JavaScript
        New-HTMLTabStyle -BorderRadius 0px -RemoveShadow
        foreach ($Group in $GroupName) {
            try {
                $ADGroup = Get-WinADGroupMember -Group $Group -All -AddSelf -CountMembers
            } catch {
                Write-Warning "Show-GroupMember - Error processing group $Group. Skipping. Needs investigation why it failed."
                continue
            }
            if ($ADGroup) {
                $GroupName = $ADGroup[0].GroupName
                $DataTableID = Get-RandomStringName -Size 15 -ToLower
                New-HTMLTab -TabName $GroupName {
                    New-HTMLTab -TabName 'Default' {
                        New-HTMLSection -Title "Group membership diagram $GroupName" {
                            New-HTMLGroupDiagramDefault -ADGroup $ADGroup -RemoveAppliesTo $RemoveAppliesTo -RemoveUsers:$RemoveUsers -RemoveComputers:$RemoveComputeres -RemoveOther:$RemoveOther
                        }
                    }
                    New-HTMLTab -TabName 'Hierarchical' {
                        New-HTMLSection -Title "Group membership diagram $GroupName" {
                            New-HTMLGroupDiagramHierachical -ADGroup $ADGroup -RemoveAppliesTo $RemoveAppliesTo -RemoveUsers:$RemoveUsers -RemoveComputers:$RemoveComputeres -RemoveOther:$RemoveOther
                        }
                        New-HTMLSection -Title "Group membership table $GroupName" {
                            New-HTMLTable -DataTable $ADGroup -Filtering -DataTableID $DataTableID
                        }
                    }
                }
            }
        }
    } -Online -FilePath $FilePath -ShowHTML
}