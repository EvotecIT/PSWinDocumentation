function Get-WinADDomainComputersFullList {
    [cmdletbinding()]
    param(
        [string] $Domain,
        [Array] $ForestSchemaComputers
    )
    Write-Verbose "Getting domain information - $Domain DomainComputersFullList"
    $TimeUsers = Start-TimeLog

    if ($Extended) {
        [string] $Properties = '*'
    } else {
        [string[]] $Properties = @(
            'SamAccountName', 'Enabled', 'OperatingSystem',
            'PasswordLastSet', 'IPv4Address', 'IPv6Address', 'Name', 'DNSHostName',
            'ManagedBy', 'OperatingSystemVersion', 'OperatingSystemHotfix',
            'OperatingSystemServicePack' , 'PasswordNeverExpires',
            'PasswordNotRequired', 'UserPrincipalName',
            'LastLogonDate', 'LockedOut', 'LogonCount',
            'CanonicalName', 'SID', 'Created', 'Modified',
            'Deleted', 'MemberOf'
            if ($ForestSchemaComputers.Name -contains 'ms-Mcs-AdmPwd') {
                'ms-Mcs-AdmPwd'
                'ms-Mcs-AdmPwdExpirationTime'
            }
        )
    }
    [string[]] $ExcludeProperty = '*Certificate', 'PropertyNames', '*Properties', 'PropertyCount', 'Certificates', 'nTSecurityDescriptor'

    Get-ADComputer -Server $Domain -Filter * -ResultPageSize 500000 -Properties $Properties -ErrorAction SilentlyContinue #| Select-Object -Property $Properties -ExcludeProperty $ExcludeProperty

    $EndUsers = Stop-TimeLog -Time $TimeUsers -Option OneLiner
    Write-Verbose "Getting domain information - $Domain DomainComputersFullList Time: $EndUsers"
}