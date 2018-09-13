function Get-AWSIAMDetails {
    [CmdletBinding()]
    param (
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $IAMDetailsList = New-Object System.Collections.ArrayList
    $IAMUsers = Get-IAMUsers -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion

    foreach ($user in $IAMUsers) {

        $groupsTemp = (Get-IAMGroupForUser -UserName $user.UserName -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion).GroupName
        $mfaTemp = (Get-IAMMFADevice -UserName $user.UserName -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion).EnableDate
        $accessKeysCreationDateTemp = (Get-IAMAccessKey -UserName $user.UserName -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion).CreateDate

        $IAMUser = [ordered] @{
            UserName              = $user.UserName
            UserGroups            = if ([string]::IsNullOrEmpty($groupsTemp)) { "No groups assigned" } Else { $groupsTemp -join ", " }
            MFAEnabledSince       = if ([string]::IsNullOrEmpty($mfaTemp)) { "Missing MFA" } Else { $mfaTemp }
            AccessKeysCount       = $accessKeysCreationDateTemp.Count
            AccessKeyCreationDate = $accessKeysCreationDateTemp -join ", "
        }
        [void]$IAMDetailsList.Add($IAMUser)
    }
    return Format-TransposeTable $IAMDetailsList
}
