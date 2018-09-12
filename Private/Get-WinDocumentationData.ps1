function Get-WinDocumentationData {
    param (
        [Object] $Data,
        [hashtable] $Forest,
        [string] $Domain
    )
    if ($Data -ne $null) {
        $Type = Get-ObjectType -Object $Data -ObjectName 'Get-WinDocumentationData' #-Verbose
        #Write-Verbose "Get-WinDocumentationData - Type: $($Type.ObjectTypeName) - Tabl $Data"
        if ($Type.ObjectTypeName -eq 'ActiveDirectory' -and $Data.ToString() -like 'Forest*') {
            return $Forest."$Data"
        } elseif ($Type.ObjectTypeName -eq 'ActiveDirectory' -and $Data.ToString() -like 'Domain*' ) {
            return $Forest.FoundDomains.$Domain."$Data"
        }
    }
    #Write-Verbose 'Get-WinDocumentationData - Data was $null'
    return
}
