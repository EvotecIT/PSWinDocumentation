function Get-WinADForestOptionalFeatures {
    [CmdletBinding()]
    param(

    )
    $OptionalFeatures = $(Get-ADOptionalFeature -Filter * )
    $Optional = [ordered]@{
        'Recycle Bin Enabled'                          = 'N/A'
        'Privileged Access Management Feature Enabled' = 'N/A'
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