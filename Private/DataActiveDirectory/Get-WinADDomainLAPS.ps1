function Get-WinADDomainLAPS {
    [CmdletBinding()]
    param(
        [string] $Domain,
        [Array] $Computers
    )
    $Properties = @(
        'Name',
        'OperatingSystem',
        'DistinguishedName',
        'ms-Mcs-AdmPwd',
        'ms-Mcs-AdmPwdExpirationTime'
    )
    [DateTime] $CurrentDate = Get-Date

    if ($null -eq $Computers -or $Computers.Count -eq 0) {
        $Computers = Get-ADComputer -Filter * -Properties $Properties
    }
    foreach ($Computer in $Computers) {
        [PSCustomObject] @{
            'Name'                 = $Computer.Name
            'Operating System'     = $Computer.'OperatingSystem'
            'Laps Password'        = $Computer.'ms-Mcs-AdmPwd'
            'Laps Expire (days)'   = Convert-TimeToDays -StartTime ($CurrentDate) -EndTime (Convert-ToDateTime -Timestring ($Computer.'ms-Mcs-AdmPwdExpirationTime'))
            'Laps Expiration Time' = Convert-ToDateTime -Timestring ($Computer.'ms-Mcs-AdmPwdExpirationTime')
            'DistinguishedName'    = $Computer.'DistinguishedName'
        }
    }
}

