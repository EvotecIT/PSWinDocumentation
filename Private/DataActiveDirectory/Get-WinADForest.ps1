function Get-WinForest {
    param(
    )
    return (Get-ADForest | Select-Object -Property * -ExcludeProperty PropertyNames, AddedProperties, RemovedProperties, ModifiedProperties, PropertyCount)
}