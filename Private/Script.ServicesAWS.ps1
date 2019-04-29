
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