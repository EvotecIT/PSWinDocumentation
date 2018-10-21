function Test-ConfigurationCredentials {
    [CmdletBinding()]
    param (
        [Object] $Configuration,
        $AllowEmptyKeys
    )
    foreach ($Key in $Configuration.Keys) {
        if ($AllowEmptyKeys -notcontains $Key -and [string]::IsNullOrWhiteSpace($Configuration.$Key)) {
            Write-Verbose "Test-ConfigurationCredentials - Configuration $Key is Null or Empty! Terminating"
            return $false
        }
    }
    return $true
}