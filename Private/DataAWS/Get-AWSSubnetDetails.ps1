function Get-AWSSubnetDetails {
    [CmdletBinding()]
    param (
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $NetworkDetailsList = New-Object System.Collections.ArrayList
    try {
        $Subnets = Get-EC2Subnet -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion
    } catch {
        $ErrorMessage = $_.Exception.Message
        Write-Warning "Get-AWSSubnetDetails - Error: $ErrorMessage"
        return
    }

    foreach ($subnet in $Subnets) {
        $SN = [pscustomobject] @{
            "Subnet ID"    = $subnet.SubnetId
            "Subnet Name"  = $subnet.Tags | Where-Object {$_.key -eq "Name"} | Select-Object -Expand Value
            "CIDR"         = $subnet.CidrBlock
            "Available IP" = $subnet.AvailableIpAddressCount
            "VPC"          = (Get-EC2Vpc -VpcId $subnet.VpcId -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion).Tags | Where-Object {$_.key -eq "Name"} | Select-Object -Expand Value
        }
        [void]$NetworkDetailsList.Add($SN)
    }
    return $NetworkDetailsList
}