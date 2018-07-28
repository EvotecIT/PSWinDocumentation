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
    $OptionalFeatures = $(Get-ADOptionalFeature -Filter * )
    $Data.OptionalFeatures = [ordered] @{
        'Recycle Bin Enabled'                          = ''
        #'Recycle Bin Scopes' = ''
        'Privileged Access Management Feature Enabled' = ''
        #'Privileged Access Management Feature Scopes' ''
    }
    ### Fix Optional Features
    foreach ($Feature in $OptionalFeatures) {
        if ($Feature.Name -eq 'Recycle Bin Feature') {
            if ("$($Feature.EnabledScopes)" -eq '') {
                $Data.OptionalFeatures.'Recycle Bin Enabled' = $False
            } else {
                $Data.OptionalFeatures.'Recycle Bin Enabled' = $True
            }
        }
        if ($Feature.Name -eq 'Privileged Access Management Feature') {
            if ("$($Feature.EnabledScopes)" -eq '') {
                $Data.OptionalFeatures.'Privileged Access Management Feature Enabled' = $False
            } else {
                $Data.OptionalFeatures.'Privileged Access Management Feature Enabled' = $True
            }
        }
    }
    ### Fix optional features
    return $Data
}