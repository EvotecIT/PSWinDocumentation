function Get-WinADForest {
    $ForestInformation = $(Get-ADForest)
    return $ForestInformation
}

function Get-WinADForestInformation {
    param (
        $ForestInformation
    )
    $Data = @{}
    $Data.ForestInformation = [ordered] @{
        'Name'                    = $ForestInformation.Name
        'Root Domain'             = $ForestInformation.RootDomain
        'Forest Functional Level' = $ForestInformation.ForestMode
        'Domains Count'           = ($ForestInformation.Domains).Count
        'Sites Count'             = ($ForestInformation.Sites).Count
        'Domains'                 = ($ForestInformation.Domains) -join ", "
        'Sites'                   = ($ForestInformation.Sites) -join ", "
    }
    $UPNSuffixList = @()
    $UPNSuffixList += $ForestInformation.RootDomain + ' (Primary/Default UPN)'
    $UPNSuffixList += $ForestInformation.UPNSuffixes
    $Data.UPNSuffixes = $UPNSuffixList
    $Data.GlobalCatalogs = $ForestInformation.GlobalCatalogs
    $Data.SPNSuffixes = $ForestInformation.SPNSuffixes
    $Data.FSMO = [ordered] @{
        'Domain Naming Master' = $ForestInformation.DomainNamingMaster
        'Schema Master'        = $ForestInformation.SchemaMaster
    }
    return $Data
}