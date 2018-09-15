function Initialize-ConfigurationCredentials {
    [CmdletBinding()]
    param (
        $Configuration
    )
    foreach ($Key in $Configuration.Keys) {
        if ([string]::IsNullOrWhiteSpace($Configuration.$Key)) {
            Write-Verbose "Verify-Credentials - Configuration $Key is Null or Empty! Terminating"
            return $false
        }
    }
    return $true
}