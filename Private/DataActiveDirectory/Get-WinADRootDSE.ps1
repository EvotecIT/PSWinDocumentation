function Get-WinADRootDSE {
    param(

    )
    return Get-ADRootDSE -Properties * | Select-Object -Property * -ExcludeProperty PropertyNames, AddedProperties, RemovedProperties, ModifiedProperties, PropertyCount
}