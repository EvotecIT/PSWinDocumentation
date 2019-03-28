function Get-WinADDomainBitlocker {
    param(
        [string] $Domain,
        [Array] $Computers
    )
    $Properties = @(
        'Name',
        'OperatingSystem',
        'DistinguishedName'
    )
    [DateTime] $CurrentDate = Get-Date

    if ($null -eq $Computers) {
        $Computers = Get-ADComputer -Filter * -Properties $Properties
    }
    foreach ($Computer in $Computers) {
        $Bitlockers = Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -SearchBase $Computer.DistinguishedName -Properties 'WhenCreated', 'msFVE-RecoveryPassword' #|  Sort-Object whenCreated -Descending #| Select-Object whenCreated, msFVE-RecoveryPassword
        foreach ($Bitlocker in $Bitlockers) {
            [PSCustomObject] @{
                'Name'                        = $Computer.Name
                'Operating System'            = $Computer.'OperatingSystem'
                'Bitlocker Recovery Password' = $Bitlocker.'msFVE-RecoveryPassword'
                'Bitlocker When'              = $Bitlocker.WhenCreated
                'DistinguishedName'           = $Computer.'DistinguishedName'
            }
        }
    }
}