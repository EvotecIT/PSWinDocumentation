function Get-ReportO365TenantDomains {
    param(

    )
    $O365UAzureTenantDomains = Get-MsolDomain | Select-Object Authentication, Capabilities, IsDefault, IsInitial, Name, RootDomain, Status, VerificationMethod   
    foreach ($Domain in $O365UAzureTenantDomains) {
        [PsCustomObject] @{
            'Domain Name'         = $Domain.Name
            'Default'             = $Domain.IsDefault
            'Initial'             = $Domain.IsInitial
            'Status'              = $Domain.Status
            'Verification Method' = $Domain.VerificationMethod
            'Capabilities'        = $Domain.Capabilities
            'Authentication'      = $Domain.Authentication
        }
    }
}