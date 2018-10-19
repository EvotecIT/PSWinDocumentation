Import-Module PSWinDocumentation -Force
Import-Module PSSharedGoods -Force

$session = Connect-WinExchange -ConnectionURI 'http://ex2013x3.ad.evotec.xyz/Powershell' -SessionName 'Evotec' -Authentication 'Kerberos' -Verbose
$Test = Import-PSSession -Session $Session -AllowClobber -DisableNameChecking

$Exchange = Get-WinExchangeInformation -Verbose
$Exchange | ft -a