function Get-AWSSubnetDetails {
    [CmdletBinding()]
    param (
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $NetworkDetailsList = New-Object System.Collections.ArrayList
    $Subnets = Get-EC2Subnet -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion

    foreach ($subnet in $Subnets) {
        $SN = [ordered] @{
            SubnetId    = $subnet.SubnetId
            SubnetName  = $subnet.Tags | Where-Object {$_.key -eq "Name"} | Select-Object -Expand Value
            CIDR        = $subnet.CidrBlock
            AvailableIp = $subnet.AvailableIpAddressCount
            VPC         = (Get-EC2Vpc -VpcId $subnet.VpcId -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion).Tags | Where-Object {$_.key -eq "Name"} | Select-Object -Expand Value
        }
        [void]$NetworkDetailsList.Add($SN)
    }
    return Format-TransposeTable $NetworkDetailsList
}