function Start-Documentation {
    [CmdletBinding()]
    param (
        [System.Collections.IDictionary] $Document
    )
    $TimeTotal = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    Test-Configuration -Document $Document

    if ($Document.DocumentAD.Enable) {

        if ($null -eq $Document.DocumentAD.Services) {
            $Document.DocumentAD.Services = ($Script:Services).Clone()
            $Document.DocumentAD.Services.OnPremises.ActiveDirectory.PasswordTests = @{
                Use                       = $Document.DocumentAD.Configuration.PasswordTests.Use
                PasswordFilePathClearText = $Document.DocumentAD.Configuration.PasswordTests.PasswordFilePathClearText
                # Fair warning it will take ages if you use HaveIBeenPwned DB :-)
                UseHashDB                 = $Document.DocumentAD.Configuration.PasswordTests.UseHashDB
                PasswordFilePathHash      = $Document.DocumentAD.Configuration.PasswordTests.PasswordFilePathHash
            }
        }
        Start-DocumentationAD -Document $Document
    }
    if ($Document.DocumentAWS.Enable) {

        if ($null -eq $Document.DocumentAWS.Services) {
            $Document.DocumentAWS.Services = ($Script:ServicesAWS).Clone()
            $Document.DocumentAWS.Services.Amazon.Credentials.AccessKey = $Document.DocumentAWS.Configuration.AWSAccessKey
            $Document.DocumentAWS.Services.Amazon.Credentials.SecretKey = $Document.DocumentAWS.Configuration.AWSSecretKey
            $Document.DocumentAWS.Services.Amazon.Credentials.Region = $Document.DocumentAWS.Configuration.AWSRegion
        }
        Start-DocumentationAWS -Document $Document
    }
    #if ($Document.DocumentExchange.Enable) {
    #    Start-DocumentationExchange -Document $Document
    #}
    if ($Document.DocumentOffice365.Enable) {
        if ($null -eq $Document.DocumentOffice365.Services) {
            $Document.DocumentOffice365.Services = ($Script:ServicesO365).Clone()

            $Document.DocumentOffice365.Services.Office365.Credentials = [ordered] @{
                Username         = $Document.DocumentOffice365.Configuration.O365Username
                Password         = $Document.DocumentOffice365.Configuration.O365Password
                PasswordAsSecure = $Document.DocumentOffice365.Configuration.O365PasswordAsSecure
                PasswordFromFile = $Document.DocumentOffice365.Configuration.O365PasswordFromFile
            }

            $Document.DocumentOffice365.Services.Office365.Azure.Use = $Document.DocumentOffice365.Configuration.O365AzureADUse
            $Document.DocumentOffice365.Services.Office365.Azure.Prefix = ''
            $Document.DocumentOffice365.Services.Office365.Azure.SessionName = 'O365Azure' # MSOL
            $Document.DocumentOffice365.Services.Office365.AzureAD.Use = $Document.DocumentOffice365.Configuration.O365AzureADUse
            $Document.DocumentOffice365.Services.Office365.AzureAD.SessionName = 'O365AzureAD' # Azure
            $Document.DocumentOffice365.Services.Office365.AzureAD.Prefix = ''

            $Document.DocumentOffice365.Services.Office365.ExchangeOnline.Use = $Document.DocumentOffice365.Configuration.O365ExchangeUse

            $Document.DocumentOffice365.Services.Office365.ExchangeOnline.Authentication = $Document.DocumentOffice365.Configuration.O365ExchangeAuthentication
            $Document.DocumentOffice365.Services.Office365.ExchangeOnline.ConnectionURI = $Document.DocumentOffice365.Configuration.O365ExchangeURI
            $Document.DocumentOffice365.Services.Office365.ExchangeOnline.Prefix = ''
            $Document.DocumentOffice365.Services.Office365.ExchangeOnline.SessionName = $Document.DocumentOffice365.Configuration.O365ExchangeSessionName

        }
        Start-DocumentationO365 -Document $Document
    }
    $TimeTotal.Stop()
    Write-Verbose "Time total: $($TimeTotal.Elapsed)"
}