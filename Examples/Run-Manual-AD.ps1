Clear-Host
Import-Module PSWinDocumentation -Force
Import-Module PSSharedGoods
Import-Module ActiveDirectory

$Forest = Get-WinADForestInformation -Verbose
$Forest.FoundDomains.'ad.evotec.pl'
$Forest.FoundDomains.'ad.evotec.pl'.DomainFineGrainedPoliciesUsers | Format-Table -AutoSize
$Forest.FoundDomains.'ad.evotec.xyz'.DomainRIDs | Format-Table -AutoSize

$Domain = Get-WinADDomainInformation -Domain 'ad.evotec.pl' -Verbose
$Domain.DomainFineGrainedPolicies
$Domain.DomainFineGrainedPoliciesUsers | Format-Table -AutoSize
$Domain.DomainFineGrainedPoliciesUsersExtended | format-Table -AutoSize