function Start-Documentation {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    #$TimeTotal = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    Test-Configuration -Document $Document

    if ($Document.DocumentAD.Enable) {

        if ($null -eq $Document.DocumentAD.Services) {
            $Document.DocumentAD.Services = ($Script:Services).Clone()
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
    if ($Document.DocumentExchange.Enable) {
        Start-DocumentationExchange -Document $Document
    }
    if ($Document.DocumentOffice365.Enable) {
        Start-DocumentationO365 -Document $Document
    }
}