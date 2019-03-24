function Get-WinADDomainOrganizationalUnitsACL {
    [cmdletbinding()]
    param(
        $DomainOrganizationalUnitsClean,
        [string] $Domain,
        [string] $NetBiosName,
        [string] $RootDomainNamingContext
    )
    Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsBasicACL"
    $Time = Start-TimeLog
    $OUs = @(
        #@{ Name = 'Root'; Value = $RootDomainNamingContext }
        foreach ($OU in $DomainOrganizationalUnitsClean) {
            @{ Name = 'Organizational Unit'; Value = $OU.DistinguishedName }
        }
    )
    $null = New-PSDrive -Name $NetBiosName -Root '' -PsProvider ActiveDirectory -Server $Domain

    @(
        foreach ($OU in $OUs) {
            $ACL = Get-Acl -Path "$NetBiosName`:\$($OU.Value)"
            [PsCustomObject] @{
                'Distinguished Name'        = $OU.Value
                'Type'                      = $OU.Name
                'Owner'                     = $ACL.Owner
                'Group'                     = $ACL.Group
                'Are AccessRules Protected' = $ACL.AreAccessRulesProtected
                'Are AuditRules Protected'  = $ACL.AreAuditRulesProtected
                'Are AccessRules Canonical' = $ACL.AreAccessRulesCanonical
                'Are AuditRules Canonical'  = $ACL.AreAuditRulesCanonical
                #'Sddl'                      = $ACL.Sddl
            }
            <#
            Get-Acl -Path "$NetBiosName`:\$($OU.Value)" | Select-Object `
            @{name = 'Distinguished Name'; expression = { $OU.Value } },
            @{name = 'Type'; expression = { $OU.Name } },
            @{name = 'Owner'; expression = { $_.Owner } },
            @{name = 'Group'; expression = { $_.Group } },
            @{name = 'Are AccessRules Protected'; expression = { $_.AreAccessRulesProtected } },
            @{name = 'Are AuditRules Protected'; expression = { $_.AreAuditRulesProtected } },
            @{name = 'Are AccessRules Canonical'; expression = { $_.AreAccessRulesCanonical } },
            @{name = 'Are AuditRules Canonical'; expression = { $_.AreAuditRulesCanonical } },
            @{name = 'Sddl'; expression = { $_.Sddl } }
            #>
        }
    )
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsBasicACL Time: $EndTime"
}
