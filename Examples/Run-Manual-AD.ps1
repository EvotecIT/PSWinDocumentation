Import-Module PSWriteWord
Import-Module PSWriteExcel
Import-Module PSWinDocumentation
Import-Module PSWriteColor
Import-Module ActiveDirectory

$Forest = Get-WinADForestInformation -Verbose
$Forest
#$Forest.FoundDomains
#$Domain = Get-WinADDomainInformation -Domain 'ad.evotec.xyz'
#$Domain