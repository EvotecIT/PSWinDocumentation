function Get-WinADDomainGUIDs {
    [cmdletbinding()]
    param(
        [string] $Domain,
        [Microsoft.ActiveDirectory.Management.ADEntity] $RootDSE
    )
    $Time = Start-TimeLog
    if ($null -eq $RootDSE) {
        $RootDSE = Get-ADRootDSE -Server $Domain
    }
    Write-Verbose "Getting domain information - $Domain DomainGUIDS"
    <#
    $GUID = @{ }
    Get-ADObject -SearchBase $RootDSE.schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID | ForEach-Object {
        if ($GUID.Keys -notcontains $_.schemaIDGUID ) {
            $GUID.add([System.GUID]$_.schemaIDGUID, $_.name)
        }
    }
    Get-ADObject -SearchBase "CN=Extended-Rights,$($RootDSE.configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID | ForEach-Object {
        if ($GUID.Keys -notcontains $_.rightsGUID ) {
            $GUID.add([System.GUID]$_.rightsGUID, $_.name)
        }
    }

#>
    $GUID = @{ }
    $Schema = Get-ADObject -SearchBase $RootDSE.schemaNamingContext -LDAPFilter '(schemaIDGUID=*)' -Properties name, schemaIDGUID
    foreach ($S in $Schema) {
        if ($GUID.Keys -notcontains $S.schemaIDGUID ) {
            $GUID.add([System.GUID]$S.schemaIDGUID, $S.name)
        }
    }

    $Extended = Get-ADObject -SearchBase "CN=Extended-Rights,$($RootDSE.configurationNamingContext)" -LDAPFilter '(objectClass=controlAccessRight)' -Properties name, rightsGUID
    foreach ($S in $Extended) {
        if ($GUID.Keys -notcontains $S.rightsGUID ) {
            $GUID.add([System.GUID]$S.rightsGUID, $S.name)
        }
    }
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose "Getting domain information - $Domain DomainGUIDS Time: $EndTime"
    return $GUID
}