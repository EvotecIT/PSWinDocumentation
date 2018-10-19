Clear-Host
Import-Module PSWinDocumentation #-Force
Import-Module PSSharedGoods
Import-Module ActiveDirectory
Import-Module PSWinDocumentation -Force

$Forest = Get-WinADForestInformation -Verbose
#$Forest.FoundDomains.'ad.evotec.pl'
#$Forest.FoundDomains.'ad.evotec.pl'.DomainFineGrainedPoliciesUsers | Format-Table -AutoSize
#$Forest.FoundDomains.'ad.evotec.xyz'.DomainRIDs | Format-Table -AutoSize
$Forest.FoundDomains.'ad.evotec.xyz'.DomainUsers
return
$User = $Forest.FoundDomains.'ad.evotec.xyz'.DomainUsers[20] | Select *
$User | Select DisplayName, PasswordLastSet


$PasswordDaysSinceChange = $User.PasswordLastSet - [DateTime]::Today
$PasswordDaysSinceChange

return
$Domain = Get-WinADDomainInformation -Domain 'ad.evotec.pl' -Verbose
$Domain.DomainFineGrainedPolicies
$Domain.DomainFineGrainedPoliciesUsers | Format-Table -AutoSize
$Domain.DomainFineGrainedPoliciesUsersExtended | format-Table -AutoSize