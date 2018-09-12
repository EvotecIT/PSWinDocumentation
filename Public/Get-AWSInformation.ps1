function Get-AWSInformation {
    param(
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $Data = [ordered] @{}
    $Data.AWSEC2Details = Get-AWSEC2Details -AWSAccessKey $AWSAccessKey -AWSSecretKey $AWSSecretKey -AWSRegion $AWSRegion
    $Data.AWESRDSDetails = Get-AWSRDSDetails -AWSAccessKey $AWSAccessKey -AWSSecretKey $AWSSecretKey -AWSRegion $AWSRegion
    $Data.AWSLBDetails = Get-AWSLBDetails -AWSAccessKey $AWSAccessKey -AWSSecretKey $AWSSecretKey -AWSRegion $AWSRegion
    $Data.AWSNetworkDetails = Get-AWSSubnetDetails -AWSAccessKey $AWSAccessKey -AWSSecretKey $AWSSecretKey -AWSRegion $AWSRegion
    $Data.AWSEIPDetails = Get-AWSElasticIpDetails -AWSAccessKey $AWSAccessKey -AWSSecretKey $AWSSecretKey -AWSRegion $AWSRegion
    $Data.AWSIAMDetails = Get-AWSIAMDetails -AWSAccessKey $AWSAccessKey -AWSSecretKey $AWSSecretKey -AWSRegion $AWSRegion

    return $Data
}