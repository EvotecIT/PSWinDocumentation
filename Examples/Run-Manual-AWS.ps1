Import-Module AWSPowerShell
Import-Module PSSharedGoods
Import-Module PSWinDocumentation

$Configuration = [ordered] @{
    AWSAccessKey = ''
    AWSSecretKey = ''
    AWSRegion    = ''
}

$AWS = Get-AWSInformation -AWSAccessKey $Configuration.AWSAccessKey -AWSSecretKey $Configuration.AWSSecretKey -AWSRegion $Configuration.AWSRegion
$AWS.AWSEC2Details | Format-Table -AutoSize
$AWS.AWSRDSDetails | Format-Table -AutoSize
$AWS.AWSLBDetails | Format-Table -AutoSize
$AWS.AWSNetworkDetails | Format-Table -AutoSize
$AWS.AWSEIPDetails | Format-Table -AutoSize
$AWS.AWSIAMDetails | Format-Table -AutoSize
