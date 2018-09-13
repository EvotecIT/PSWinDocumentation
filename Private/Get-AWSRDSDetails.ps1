function Get-AWSRDSDetails {
    [CmdletBinding()]
    param (
        [string] $AWSAccessKey,
        [string] $AWSSecretKey,
        [string] $AWSRegion
    )

    $RDSDetailsList = New-Object System.Collections.ArrayList
    $RDSInstances = Get-RDSDBInstance -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion

    foreach ($instance in $RDSInstances) {
        $RDS = [ordered] @{
            InstanceName          = $instance.DBInstanceIdentifier
            InstanceClass         = $instance.DBInstanceClass
            MutliAz               = if ($instance.Engine.StartsWith("aurora")) { "not applicable" } Else { $instance.MultiAz }
            InstanceEngine        = $instance.Engine
            InstanceEngineVersion = $instance.EngineVersion
            Storage               = if ($instance.Engine.StartsWith("aurora")) { "Dynamic" } Else { [string]::Format("{0} GB", $instance.AllocatedStorage) }
            Environment           = Get-RDSTagForResource -AccessKey $AWSAccessKey -SecretKey $AWSSecretKey -Region $AWSRegion -ResourceName $instance.DBInstanceArn | Where-Object {$_.key -eq "Environment"} | Select-Object -Expand Value

        }
        [void]$RDSDetailsList.Add($RDS)
    }
    return Format-TransposeTable $RDSDetailsList
}
