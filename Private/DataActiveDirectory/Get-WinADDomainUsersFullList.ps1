function Get-WinADDomainUsersFullList {
    [CmdletBinding()]
    param(
        [string] $Domain,
        [switch] $Extended,
        [Array] $ForestSchemaUsers
    )
    Write-Verbose "Getting domain information - $Domain DomainUsersFullList"
    $TimeUsers = Start-TimeLog
    if ($Extended) {
        [string] $Properties = '*'
    } else {
        $Properties = @(
            'Name'
            'UserPrincipalName'
            'SamAccountName'
            'DisplayName'
            'GivenName'
            'Surname'
            'EmailAddress'
            'PasswordExpired'
            'PasswordLastSet'
            'PasswordNotRequired'
            'PasswordNeverExpires'
            'Enabled'
            'Manager'
            'msDS-UserPasswordExpiryTimeComputed'
            'AccountExpirationDate'
            'AccountLockoutTime'
            'AllowReversiblePasswordEncryption'
            'BadLogonCount'
            'CannotChangePassword'
            'CanonicalName'
            'Description'
            'DistinguishedName'
            'EmployeeID'
            'EmployeeNumber'
            'LastBadPasswordAttempt'
            'LastLogonDate'
            'Created'
            'Modified'
            'ProtectedFromAccidentalDeletion'
            'PrimaryGroup'
            'MemberOf'
            if ($ForestSchemaUsers.Name -contains 'ExtensionAttribute1') {
                'ExtensionAttribute1'
                'ExtensionAttribute2'
                'ExtensionAttribute3'
                'ExtensionAttribute4'
                'ExtensionAttribute5'
                'ExtensionAttribute6'
                'ExtensionAttribute7'
                'ExtensionAttribute8'
                'ExtensionAttribute9'
                'ExtensionAttribute10'
                'ExtensionAttribute11'
                'ExtensionAttribute12'
                'ExtensionAttribute13'
                'ExtensionAttribute14'
                'ExtensionAttribute15'
            }
        )
    }

    [string[]] $ExcludeProperty = '*Certificate', 'PropertyNames', '*Properties', 'PropertyCount', 'Certificates', 'nTSecurityDescriptor'

    Get-ADUser -Server $Domain -ResultPageSize 500000 -Filter * -Properties $Properties #| Select-Object -Property $Properties -ExcludeProperty $ExcludeProperty

    $EndUsers = Stop-TimeLog -Time $TimeUsers -Option OneLiner
    Write-Verbose "Getting domain information - $Domain DomainUsersFullList Time: $EndUsers"
}