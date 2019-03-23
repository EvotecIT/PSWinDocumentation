function Get-WinADDomainOrganizationalUnits {
    [CmdletBinding()]
    param(
        [string] $Domain,
        [Array] $OrgnaizationalUnits
    )
    Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnits"
    if ($null -eq $OrgnaizationalUnits) {
        $OrgnaizationalUnits = $(Get-ADOrganizationalUnit -Server $Domain -Properties * -Filter * )
    }
    $TimeOU = Start-TimeLog
    $Output = foreach ($O in $OrgnaizationalUnits) {
        [PSCustomObject] @{
            'Canonical Name'  = $O.CanonicalName
            'Managed By'      = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $_.ManagedBy -Verbose).Name
            'Manager Email'   = (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $_.ManagedBy -Verbose).EmailAddress
            'Protected'       = $O.ProtectedFromAccidentalDeletion
            Description       = $O.Description
            Created           = $O.Created
            Modified          = $O.Modified
            Deleted           = $O.Deleted
            'Postal Code'     = $O.PostalCode
            City              = $O.City
            Country           = $O.Country
            State             = $O.State
            'Street Address'  = $O.StreetAddress
            DistinguishedName = $O.DistinguishedName
            ObjectGUID        = $O.ObjectGUID
        }
    }
    $Output | Sort-Object 'Canonical Name'
    $EndOU = Stop-TimeLog -Time $TimeOU -Option OneLiner
    Write-Verbose -Message "Getting domain information - $Domain DomainOrganizationalUnits Time: $EndOU"
    <#
        $Time44 = Start-TimeLog
    for ($i = 1; $i -lt 1000; $i++) {
        $OrgnaizationalUnits | Select-Object `
        @{ n = 'Canonical Name'; e = { $_.CanonicalName } },
        @{ n = 'Managed By'; e = {
                (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $_.ManagedBy -Verbose).Name
            }
        },
        @{ n = 'Manager Email'; e = {
                (Get-ADObjectFromDistingusishedName -ADCatalog $Data.DomainUsersFullList -DistinguishedName $_.ManagedBy -Verbose).EmailAddress
            }
        },
        @{ n = 'Protected'; e = { $_.ProtectedFromAccidentalDeletion } },
        Created,
        Modified,
        Deleted,
        @{ n = 'Postal Code'; e = { $_.PostalCode } },
        City,
        Country,
        State,
        @{ n = 'Street Address'; e = { $_.StreetAddress } },
        DistinguishedName,
        ObjectGUID | Sort-Object 'Canonical Name'

    }
    $End = Stop-TimeLog -Time $Time44 -Option OneLiner
    Write-Verbose $end
    #>
}