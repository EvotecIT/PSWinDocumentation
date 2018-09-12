function Get-WinDocumentationData {
    param (
        [alias("Data")][Object] $DataToGet,
        [alias("Forest")][Object] $Object,
        [string] $Domain
    )
    if ($DataToGet -ne $null) {
        $Type = Get-ObjectType -Object $DataToGet -ObjectName 'Get-WinDocumentationData' -Verbose
        Write-Verbose "Get-WinDocumentationData - DataToGet: $DataToGet Domain: $Domain"
        if ($Type.ObjectTypeName -eq 'ActiveDirectory') {
            if ($DataToGet.ToString() -like 'Forest*') {
                return $Object."$DataToGet"
            } elseif ($DataToGet.ToString() -like 'Domain*' ) {
                return $Object.FoundDomains.$Domain."$DataToGet"
            }
        } else {

            return $Object.$DataToGet
        }
    }
    return
}
