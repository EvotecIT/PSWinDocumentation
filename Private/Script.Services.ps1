$Script:Services = @{
    OnPremises = [ordered] @{
        Credentials     = [ordered] @{
            Username         = ''
            Password         = ''
            PasswordAsSecure = $true
            PasswordFromFile = $true
        }
        ActiveDirectory = [ordered] @{
            Use           = $true
            OnlineMode    = $true

            Import        = @{
                Use  = $false
                From = 'Folder' # Folder
                Path = "$Env:USERPROFILE\Desktop\PSWinDocumentation"
                # or "$Env:USERPROFILE\Desktop\PSWinDocumentation\PSWinDocumentation.xml"
            }
            Export        = @{
                Use        = $false
                To         = 'Folder' # Folder/File/Both
                FolderPath = "$Env:USERPROFILE\Desktop\PSWinDocumentation"
                FilePath   = "$Env:USERPROFILE\Desktop\PSWinDocumentation\PSWinDocumentation.xml"
            }

            Prefix        = ''
            SessionName   = 'ActiveDirectory' # MSOL

            PasswordTests = @{
                Use                       = $false
                PasswordFilePathClearText = 'C:\Support\GitHub\PSWinDocumentation\Ignore\Passwords.txt'
                # Fair warning it will take ages if you use HaveIBeenPwned DB :-)
                UseHashDB                 = $false
                PasswordFilePathHash      = 'C:\Support\GitHub\PSWinDocumentation\Ignore\Passwords-Hashes.txt'
            }
        }
    }
}