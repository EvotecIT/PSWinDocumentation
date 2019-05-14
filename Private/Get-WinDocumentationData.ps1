function Get-WinDocumentationData {
    param (
        [alias("Data")][Object] $DataToGet,
        [alias("Forest")][Object] $Object,
        [string] $Domain
    )
    if ($null -ne $DataToGet) {
        $Type = Get-ObjectType -Object $DataToGet -ObjectName 'Get-WinDocumentationData' #-Verbose
        if ($Type.ObjectTypeName -eq 'ActiveDirectory') {
            #Write-Verbose "Get-WinDocumentationData - DataToGet: $DataToGet Domain: $Domain"
            if ("$DataToGet" -like 'Forest*') {
                return $Object."$DataToGet"
            } elseif ($DataToGet.ToString() -like 'Domain*' ) {
                return $Object.FoundDomains.$Domain."$DataToGet"
            }
        } else {
            #Write-Verbose "Get-WinDocumentationData - DataToGet: $DataToGet Object: $($Object.Count)"
            return $Object."$DataToGet"
        }
    }
    return
}
