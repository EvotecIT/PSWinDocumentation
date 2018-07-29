function Get-WinADForestInformation {
    [CmdletBinding()]
    $Data = @{}
    $ForestInformation = $(Get-ADForest)
    $Data.Forest = $ForestInformation
    $Data.RootDSE = $(Get-ADRootDSE -Properties *)

    $Data.ForestName = $ForestInformation.Name
    $Data.ForestNameDN = $Data.RootDSE.defaultNamingContext
    $Data.Domains = $ForestInformation.Domains
    $Data.ForestInformation = [ordered] @{
        'Name'                    = $ForestInformation.Name
        'Root Domain'             = $ForestInformation.RootDomain
        'Forest Functional Level' = $ForestInformation.ForestMode
        'Domains Count'           = ($ForestInformation.Domains).Count
        'Sites Count'             = ($ForestInformation.Sites).Count
        'Domains'                 = ($ForestInformation.Domains) -join ", "
        'Sites'                   = ($ForestInformation.Sites) -join ", "
    }
    $Data.UPNSuffixes = Invoke-Command -ScriptBlock {
        $UPNSuffixList = @()
        $UPNSuffixList += $ForestInformation.RootDomain + ' (Primary/Default UPN)'
        $UPNSuffixList += $ForestInformation.UPNSuffixes
        return $UPNSuffixList
    }
    $Data.GlobalCatalogs = $ForestInformation.GlobalCatalogs
    $Data.SPNSuffixes = $ForestInformation.SPNSuffixes
    $Data.FSMO = Invoke-Command -ScriptBlock {
        $FSMO = [ordered] @{
            'Domain Naming Master' = $ForestInformation.DomainNamingMaster
            'Schema Master'        = $ForestInformation.SchemaMaster
        }
        return $FSMO
    }
    $Data.OptionalFeatures = Invoke-Command -ScriptBlock {
        $OptionalFeatures = $(Get-ADOptionalFeature -Filter * )
        $Optional = [ordered]@{
            'Recycle Bin Enabled'                          = ''
            'Privileged Access Management Feature Enabled' = ''
        }
        ### Fix Optional Features
        foreach ($Feature in $OptionalFeatures) {
            if ($Feature.Name -eq 'Recycle Bin Feature') {
                if ("$($Feature.EnabledScopes)" -eq '') {
                    $Optional.'Recycle Bin Enabled' = $False
                } else {
                    $Optional.'Recycle Bin Enabled' = $True
                }
            }
            if ($Feature.Name -eq 'Privileged Access Management Feature') {
                if ("$($Feature.EnabledScopes)" -eq '') {
                    $Optional.'Privileged Access Management Feature Enabled' = $False
                } else {
                    $Optional.'Privileged Access Management Feature Enabled' = $True
                }
            }
        }
        return $Optional
        ### Fix optional features
    }
    return $Data
}