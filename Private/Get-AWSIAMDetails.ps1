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

        $IAMUser = [pscustomobject] @{
            "User Name"                 = $user.UserName
            "Groups"                    = if ([string]::IsNullOrEmpty($groupsTemp)) { "No groups assigned" } Else { $groupsTemp -join ", " }
            "MFA Since"                 = if ([string]::IsNullOrEmpty($mfaTemp)) { "Missing MFA" } Else { $mfaTemp }
            "Access Keys Count"         = $accessKeysCreationDateTemp.Count
            "Access Keys Creation Date" = $accessKeysCreationDateTemp -join ", "
        }
        [void]$IAMDetailsList.Add($IAMUser)
    }
    return $IAMDetailsList
}
