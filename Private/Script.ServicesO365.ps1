$Script:ServicesO365 = @{
    Office365 = [ordered] @{
        Credentials    = [ordered] @{
            Username         = 'przemyslaw.klys@evotec.pl'
            Password         = 'C:\Support\Important\Password-O365-Evotec.txt'
            PasswordAsSecure = $true
            PasswordFromFile = $true
        }
        Azure          = [ordered] @{
            Use         = $true
            OnlineMode  = $true

            Import      = @{
                Use  = $false
                From = 'Folder' # Folder
                Path = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365Azure"
                # or "$Env:USERPROFILE\Desktop\PSWinDocumentation\PSWinDocumentation.xml"
            }
            Export      = @{
                Use        = $false
                To         = 'Folder' # Folder/File/Both
                FolderPath = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365Azure"
                FilePath   = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365Azure\PSWinDocumentation.xml"
            }

            ExportXML   = $false
            FilePathXML = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365Azure.xml"

            Prefix      = ''
            SessionName = 'O365Azure' # MSOL
        }
        AzureAD        = [ordered] @{
            Use         = $true
            OnlineMode  = $true

            Import      = @{
                Use  = $false
                From = 'Folder' # Folder
                Path = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365AzureAD"
                # or "$Env:USERPROFILE\Desktop\PSWinDocumentation\PSWinDocumentation.xml"
            }
            Export      = @{
                Use        = $false
                To         = 'Folder' # Folder/File/Both
                FolderPath = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365AzureAD"
                FilePath   = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365AzureAD\PSWinDocumentation.xml"
            }

            ExportXML   = $false
            FilePathXML = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365AzureAD.xml"

            SessionName = 'O365AzureAD' # Azure
            Prefix      = ''
        }
        ExchangeOnline = [ordered] @{
            Use            = $true
            OnlineMode     = $true

            Import         = @{
                Use  = $false
                From = 'Folder' # Folder
                Path = "$Env:USERPROFILE\Desktop\PSWinDocumentation"
                # or "$Env:USERPROFILE\Desktop\PSWinDocumentation\PSWinDocumentation.xml"
            }
            Export         = @{
                Use        = $false
                To         = 'Folder' # Folder/File/Both
                FolderPath = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365ExchangeOnline"
                FilePath   = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365ExchangeOnline\PSWinDocumentation.xml"
            }

            ExportXML      = $false
            FilePathXML    = "$Env:USERPROFILE\Desktop\PSWinDocumentation-O365ExchangeOnline.xml"

            Authentication = 'Basic'
            ConnectionURI  = 'https://outlook.office365.com/powershell-liveid/'
            Prefix         = 'O365'
            SessionName    = 'O365Exchange'
        }
    }
}