function Get-AWSEC2Details {
    param (
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $EC2DetailsList = New-Object System.Collections.ArrayList
    $EC2Instances = Get-EC2Instance -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion

    foreach ($instance in $EC2Instances) {
        $Ec2 = [ordered] @{
            InstanceID       = $instance[0].Instances[0].InstanceId
            InstanceName     = $instance[0].Instances[0].Tags | Where-Object {$_.key -eq "Name"} | Select-Object -Expand Value
            Environment      = $instance[0].Instances[0].Tags | Where-Object {$_.key -eq "Environment"} | Select-Object -Expand Value
            InstanceType     = $instance[0].Instances[0].InstanceType
            PrivateIpAddress = $instance[0].Instances[0].PrivateIpAddress
            PublicIpAddress  = $instance[0].Instances[0].PublicIpAddress
        }
        [void]$EC2DetailsList.Add($ec2)
    }
    return Format-TransposeTable $EC2DetailsList
}
