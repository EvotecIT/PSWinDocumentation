function Get-WinADForestSites {
    param(
        
    )
    return Get-ADReplicationSite -Filter * -Properties * | Select-Object -Property * -ExcludeProperty PropertyNames, AddedProperties, RemovedProperties, ModifiedProperties, PropertyCount
}