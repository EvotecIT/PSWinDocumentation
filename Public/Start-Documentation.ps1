function Start-Documentation {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    $TimeTotal = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
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



$Script:ServicesAWS = @{
    Amazon = [ordered] @{
        Credentials = [ordered] @{
            AccessKey = ''
            SecretKey = ''
            Region    = 'eu-west-1'
        }
        AWS         = [ordered] @{
            Use         = $true
            OnlineMode  = $true

            Import      = @{
                Use  = $false
                From = 'Folder' # Folder
                Path = "$Env:USERPROFILE\Desktop\PSWinDocumentation"
                # or "$Env:USERPROFILE\Desktop\PSWinDocumentation\PSWinDocumentation.xml"
            }
            Export      = @{
                Use        = $false
                To         = 'Folder' # Folder/File/Both
                FolderPath = "$Env:USERPROFILE\Desktop\PSWinDocumentation"
                FilePath   = "$Env:USERPROFILE\Desktop\PSWinDocumentation\PSWinDocumentation.xml"
            }

            Prefix      = ''
            SessionName = 'AWS'
        }
    }
}