function Get-AWSLBDetails {
    param (
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $LBDetailsList = New-Object System.Collections.ArrayList
    $ELBs = Get-ELBLoadBalancer -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion # Classic Load Balancers
    $ALBs = Get-ELB2LoadBalancer -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion # Application Load Balancers

    foreach ($lb in $ELBs) {
        $LB = [ordered] @{
            LBName             = $lb.LoadBalancerName
            LBType             = "ELB"
            Scheme             = $lb.Scheme
            DNSName            = $lb.DNSName
            RegistredInstances = $lb.Instances.InstanceId -join ", "
        }
        [void]$LBDetailsList.Add($LB)
    }
    foreach ($lb in $ALBs) {
        $LB = [ordered] @{
            LBName             = $lb.LoadBalancerName
            LBType             = "ALB"
            Scheme             = $lb.Scheme
            DNSName            = $lb.DNSName
            RegistredInstances = "Dynamic Routing"
        }
        [void]$LBDetailsList.Add($LB)
    }
    return Format-TransposeTable $LBDetailsList
}
