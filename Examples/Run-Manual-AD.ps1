Import-Module PSWinDocumentation
Import-Module PSSharedGoods
Import-Module ActiveDirectory

$Forest = Get-WinADForestInformation -Verbose
$Forest
#$Forest.FoundDomains
#$Domain = Get-WinADDomainInformation -Domain 'ad.evotec.xyz'
#$Domain