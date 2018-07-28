Function Get-PrivilegedGroupsMemberCount {
    Param (
        [Parameter( Mandatory = $true, ValueFromPipeline = $true )]
        $Domain,
        $DomainSID
    )

    ## Jeff W. said this was original code, but until I got ahold of it and
    ## rewrote it, it looked only slightly changed from:
    ## https://gallery.technet.microsoft.com/scriptcenter/List-Membership-In-bff89703
    ## So I give them both credit. :-)

    Write-Debug "***Get-PrivilegedGroupsMemberCount: domainName='$domainName', domainSid='$domainSidValue'"

    ## Carefully chosen from a more complete list at:
    ## https://support.microsoft.com/en-us/kb/243330
    ## Administrator (not a group, just FYI)    - $DomainSidValue-500
    ## Domain Admins                            - $DomainSidValue-512
    ## Schema Admins                            - $DomainSidValue-518
    ## Enterprise Admins                        - $DomainSidValue-519
    ## Group Policy Creator Owners              - $DomainSidValue-520
    ## BUILTIN\Administrators                   - S-1-5-32-544
    ## BUILTIN\Account Operators                - S-1-5-32-548
    ## BUILTIN\Server Operators                 - S-1-5-32-549
    ## BUILTIN\Print Operators                  - S-1-5-32-550
    ## BUILTIN\Backup Operators                 - S-1-5-32-551
    ## BUILTIN\Replicators                      - S-1-5-32-552
    ## BUILTIN\Network Configuration Operations - S-1-5-32-556
    ## BUILTIN\Incoming Forest Trust Builders   - S-1-5-32-557
    ## BUILTIN\Event Log Readers                - S-1-5-32-573
    ## BUILTIN\Hyper-V Administrators           - S-1-5-32-578
    ## BUILTIN\Remote Management Users          - S-1-5-32-580

    ## FIXME - we report on all these groups for every domain, however
    ## some of them are forest wide (thus the membership will be reported
    ## in every domain) and some of the groups only exist in the
    ## forest root.
    $PrivilegedGroups = "$DomainSID-512", "$DomainSID-518",
    "$DomainSID-519", "$DomainSID-520",
    "S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549",
    "S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552",
    "S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573",
    "S-1-5-32-578", "S-1-5-32-580"

    ForEach ( $PrivilegedGroup in $PrivilegedGroups ) {
        $source = New-Object DirectoryServices.DirectorySearcher( "LDAP://$DomainName" )
        $source.SearchScope = 'Subtree'
        $source.PageSize = 1000
        $source.Filter = "(objectSID=$PrivilegedGroup)"

        Write-Debug "***Get-PrivilegedGroupsMemberCount: LDAP://$DomainName, (objectSid=$PrivilegedGroup)"

        $Groups = $source.FindAll()
        ForEach ( $Group in $Groups ) {
            $DistinguishedName = $Group.Properties.Item( 'distinguishedName' )
            $groupName = $Group.Properties.Item( 'Name' )

            Write-Debug "***Get-PrivilegedGroupsMemberCount: searching group '$groupName'"

            $Source.Filter = "(memberOf:1.2.840.113556.1.4.1941:=$DistinguishedName)"
            $Users = $null
            ## CHECK: I don't think a try/catch is necessary here - MBS
            try {
                $Users = $Source.FindAll()
            } catch {
                # nothing
            }
            If ( $null -eq $users ) {
                ## Obsolete: F-I-X-M-E: we should probably Return a PSObject with a count of zero
                ## Write-ToCSV and Write-ToWord understand empty Return results.

                Write-Debug "***Get-PrivilegedGroupsMemberCount: no members found in $groupName"
            } Else {
                Function GetProperValue {
                    Param(
                        [Object] $object
                    )

                    If ( $object -is [System.DirectoryServices.SearchResultCollection] ) {
                        Return $object.Count
                    }
                    If ( $object -is [System.DirectoryServices.SearchResult] ) {
                        Return 1
                    }
                    If ( $object -is [Array] ) {
                        Return $object.Count
                    }
                    If ( $null -eq $object ) {
                        Return 0
                    }

                    Return 1
                }

                [int]$script:MemberCount = GetProperValue $Users

                Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' user count before first filter $MemberCount"

                $Object = New-Object -TypeName PSObject
                $Object | Add-Member -MemberType NoteProperty -Name 'Domain' -Value $Domain
                $Object | Add-Member -MemberType NoteProperty -Name 'Group'  -Value $groupName

                $Members = $Users | Where-Object { $_.Properties.Item( 'objectCategory' ).Item( 0 ) -like 'cn=person*' }
                $script:MemberCount = GetProperValue $Members

                Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' user count after first filter $MemberCount"

                Write-Debug "***Get-PrivilegedGroupsMemberCount: '$groupName' has $MemberCount members"

                $Object | Add-Member -MemberType NoteProperty -Name 'Members' -Value $MemberCount
                $Object
            }
        }
    }

}

Function Get-PrivilegedGroupsMembers {
    [CmdletBinding()]
    Param (
        $Domain,
        $DomainSID
    )
    $PrivilegedGroups1 = "$DomainSID-512", "$DomainSID-518", "$DomainSID-519", "$DomainSID-520" # will be only on root domain
    $PrivilegedGroups2 = "S-1-5-32-544", "S-1-5-32-548", "S-1-5-32-549", "S-1-5-32-550", "S-1-5-32-551", "S-1-5-32-552", "S-1-5-32-556", "S-1-5-32-557", "S-1-5-32-573", "S-1-5-32-578", "S-1-5-32-580"

    $SpecialGroups = @()
    foreach ($Group in ($PrivilegedGroups1 + $PrivilegedGroups2)) {
        Write-Verbose "Get-PrivilegedGroupsMembers - Group $Group in $Domain ($DomainSid)"
        try {
            $GroupInfo = Get-AdGroup $Group -ErrorAction Stop
            $GroupData = get-adgroupmember -Server $Domain -Identity $group | Sort-Object -Unique
            $GroupDataRecursive = get-adgroupmember -Server $Domain -Identity $group -Recursive:$Recursive | Sort-Object -Unique
            $GroupDataRecursive | fl *
            #$GroupData.SamAccountName #| Select * -Unique
            #$GroupData | ft -a
            $SpecialGroups += [ordered]@{
                'Group Name'              = $GroupInfo.Name
                'Group Category'          = $GroupInfo.GroupCategory
                'Group Scope'             = $GroupInfo.GroupScope
                'Members Count'           = Get-ObjectCount $GroupData
                'Members'                 = $GroupData.SamAccountName | Sort-Object -Unique
                'Members Recursive Count' = Get-ObjectCount $GroupDataRecursive
                'Members Recursive'       = $GroupDataRecursive.SamAccountName
            }
        } catch {}
    }
    return $SpecialGroups.ForEach( {[PSCustomObject]$_})
}
