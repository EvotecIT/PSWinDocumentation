function Get-ReportO365Subscriptions {
    param(

    )
    $O365UAzureSubscription = Get-MsolSubscription
    $Licenses = foreach ($Subscription in $O365UAzureSubscription) {
        foreach ($Plan in $Subscription.ServiceStatus) {
            [PSCustomObject] @{
                'Licenses Name'       = Convert-Office365License -License $Subscription.SkuPartNumber
                'Licenses SKU'        = $Subscription.SkuPartNumber
                'Service Plan Name'   = Convert-Office365License -License $Plan.ServicePlan.ServiceName
                'Service Plan SKU'    = $Plan.ServicePlan.ServiceName
                'Service Plan ID'     = $Plan.ServicePlan.ServicePlanId
                'Service Plan Type'   = $Plan.ServicePlan.ServiceType
                'Service Plan Class'  = $Plan.ServicePlan.TargetClass
                'Service Plan Status' = $Plan.ProvisioningStatus
                'Licenses Total'      = $Subscription.TotalLicenses
                'Licenses Status'     = $Subscription.Status
                'Licenses SKUID'      = $Subscription.SkuId
                'Licenses Are Trial'  = $Subscription.IsTrial
                'Licenses Created'    = $Subscription.DateCreated
                'Next Lifecycle Date' = $Subscription.NextLifecycleDate
                'ObjectID'            = $Subscription.ObjectId
                'Ocp SubscriptionID'  = $Subscription.OcpSubscriptionId
            }
        }
    }
    return $Licenses | Sort-Object 'Licenses Name'
}