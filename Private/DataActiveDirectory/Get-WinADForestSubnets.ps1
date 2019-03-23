function Get-WinADForestSubnets {
    [CmdletBinding()]
    param(

    )
    $ExludedProperties = @(
        'PropertyNames', 'AddedProperties', 'RemovedProperties', 'ModifiedProperties', 'PropertyCount'
    )
    $Properties = @(
        'Name', 'DisplayName', 'Description', 'Site', 'ProtectedFromAccidentalDeletion', 'Created', 'Modified', 'Deleted'
    )

    return Get-ADReplicationSubnet -Filter * -Properties $Properties | Select-Object -Property $Properties -ExcludeProperty $ExludedProperties
}