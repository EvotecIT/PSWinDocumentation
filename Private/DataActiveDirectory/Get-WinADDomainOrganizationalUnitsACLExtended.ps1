function Get-WinADDomainOrganizationalUnitsACLExtended {
    [cmdletbinding()]
    param(
        $DomainOrganizationalUnitsClean,
        [string] $Domain,
        [string] $NetBiosName,
        [string] $RootDomainNamingContext,
        $GUID
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

            $ACLs = Get-Acl -Path "$NetBiosName`:\$($OU.Value)" | Select-Object -ExpandProperty Access
            foreach ($ACL in $ACLs) {
                [PSCustomObject] @{
                    'Distinguished Name'        = $OU.Value
                    'Type'                      = $OU.Name
                    'AccessControlType'         = $ACL.AccessControlType
                    'ObjectType Name'           = if ($ACL.objectType.ToString() -eq '00000000-0000-0000-0000-000000000000') { 'All' } Else { $GUID.Item($ACL.objectType) }
                    'Inherited ObjectType Name' = $GUID.Item($ACL.inheritedObjectType)
                    'ActiveDirectoryRights'     = $ACL.ActiveDirectoryRights
                    'InheritanceType'           = $ACL.InheritanceType
                    'ObjectType'                = $ACL.ObjectType
                    'InheritedObjectType'       = $ACL.InheritedObjectType
                    'ObjectFlags'               = $ACL.ObjectFlags
                    'IdentityReference'         = $ACL.IdentityReference
                    'IsInherited'               = $ACL.IsInherited
                    'InheritanceFlags'          = $ACL.InheritanceFlags
                    'PropagationFlags'          = $ACL.PropagationFlags
                }
            }

            <#
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

            #>

        }
    )
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnitsExtended Time: $EndTime"
}

<#
$Data = @{}
$Domain = 'ad.evotec.xyz'
$Data.DomainRootDSE = $(Get-ADRootDSE -Server $Domain)
$Data.DomainInformation = $(Get-ADDomain -Server $Domain)
$Data.DomainGUIDS = Invoke-Command -ScriptBlock {
    $GUID = @{ }
    Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID | ForEach-Object {
        if ($GUID.Keys -notcontains $_.schemaIDGUID ) {
            $GUID.add([System.GUID]$_.schemaIDGUID, $_.name)
        }
    }
    Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID | ForEach-Object {
        if ($GUID.Keys -notcontains $_.rightsGUID ) {
            $GUID.add([System.GUID]$_.rightsGUID, $_.name)
        }
    }
    return $GUID
}
$OU = $(Get-ADOrganizationalUnit -Server $Domain -Properties * -Filter * )


Get-WinADDomainOrganizationalUnitsACLExtended  `
    -DomainOrganizationalUnitsClean $OU `
    -Domain $Domain `
    -NetBiosName $Data.DomainInformation.NetBIOSName `
    -RootDomainNamingContext $Data.DomainRootDSE.rootDomainNamingContext `
    -GUID $Data.DomainGUIDS

    #>