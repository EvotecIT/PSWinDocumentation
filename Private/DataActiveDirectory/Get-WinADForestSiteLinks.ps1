function Get-WinADForestSiteLinks {
    [CmdletBinding()]
    param(

    )
    $ExludedProperties = @(
        'PropertyNames', 'AddedProperties', 'RemovedProperties', 'ModifiedProperties', 'PropertyCount'
    )
    $Properties = @(
        'Name', 'Cost', 'ReplicationFrequencyInMinutes', 'ReplInterval',
        'ReplicationSchedule', 'Created', 'Modified', 'Deleted', 'InterSiteTransportProtocol',
        'DistinguishedName', 'ProtectedFromAccidentalDeletion'
        #siteList,nTSecurityDescriptor

    )

    return Get-ADReplicationSiteLink -Filter * -Properties $Properties | Select-Object -Property $Properties -ExcludeProperty $ExludedProperties
}