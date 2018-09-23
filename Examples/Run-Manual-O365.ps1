Import-Module PSWinDocumentation -Force
Import-Module PSSharedGoods -Force

$Username = 'przemyslaw.klys@evotec.pl'
$Password = ''


$session = Connect-WinExchange `
    -ConnectionURI 'https://outlook.office365.com/powershell-liveid/' `
    -SessionName 'EvotecO365' `
    -Authentication 'Basic' `
    -Username $Username `
    -Password $Password `
    -AsSecure:$False `
    -FromFile:$false `
    -Verbose

$Test = Import-PSSession -Session $Session -AllowClobber -DisableNameChecking -Prefix 'O365'

$ExchangeOnline = Get-WinO365Exchange -Verbose
$ExchangeOnline


$SessionAzure = Connect-WinAzure `
    -SessionName 'EvotecO365Azure' `
    -Username $Username `
    -Password $Password `
    -AsSecure:$False `
    -FromFile:$False

$Azure = Get-WinO365Azure -Verbose
$Azure