function Get-WinADDomainOrganizationalUnitsACLExtended {
    [cmdletbinding()]
    param(
        $DomainOrganizationalUnitsClean,
        [string] $Domain,
        [string] $NetBiosName,
        [string] $RootDomainNamingContext
    )
    Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsExtended"
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
            Get-Acl -Path "$NetBiosName`:\$($OU.Value)" | `
                Select-Object -ExpandProperty Access | `
                Select-Object `
            @{name = 'Distinguished Name'; expression = { $OU.Value } },
            @{name = 'Type'; expression = { $OU.Name } },
            @{name = 'AccessControlType'; expression = { $_.AccessControlType } },
            @{name = 'ObjectType Name'; expression = { if ($_.objectType.ToString() -eq '00000000-0000-0000-0000-000000000000') { 'All' } Else { $GUID.Item($_.objectType) } } },
            @{name = 'Inherited ObjectType Name'; expression = { $GUID.Item($_.inheritedObjectType) } },
            @{name = 'ActiveDirectoryRights'; expression = { $_.ActiveDirectoryRights } },
            @{name = 'InheritanceType'; expression = { $_.InheritanceType } },
            @{name = 'ObjectType'; expression = { $_.ObjectType } },
            @{name = 'InheritedObjectType'; expression = { $_.InheritedObjectType } },
            @{name = 'ObjectFlags'; expression = { $_.ObjectFlags } },
            @{name = 'IdentityReference'; expression = { $_.IdentityReference } },
            @{name = 'IsInherited'; expression = { $_.IsInherited } },
            @{name = 'InheritanceFlags'; expression = { $_.InheritanceFlags } },
            @{name = 'PropagationFlags'; expression = { $_.PropagationFlags } }
        }
    )
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsExtended Time: $EndTime"
}