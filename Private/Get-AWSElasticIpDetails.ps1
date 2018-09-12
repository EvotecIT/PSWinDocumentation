function Get-AWSElasticIpDetails {
    param (
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $EIPDetailsList = New-Object System.Collections.ArrayList
    $EIPs = Get-EC2Address -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion

    foreach ($eip in $EIPs) {
        $ElasticIP = [ordered] @{
            Name               = $eip.Tags | Where-Object {$_.key -eq "Name"} | Select-Object -Expand Value
            IP                 = $eip.PublicIp
            AssignedToInstance = $eip.InstanceId
            NetworkInterface   = $eip.NetworkInterfaceId

        }
        [void]$EIPDetailsList.Add($ElasticIP)
    }
    return Format-TransposeTable $EIPDetailsList
}