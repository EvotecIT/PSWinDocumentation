Clear-Host
Import-Module PSWinDocumentation -Force
Import-Module PSSharedGoods
Import-Module ActiveDirectory

#$Forest = Get-WinADForestInformation -Verbose
#$Forest.FoundDomains.'ad.evotec.pl'
#$Forest.FoundDomains.'ad.evotec.pl'.DomainFineGrainedPoliciesUsers | Format-Table -AutoSize
#$Forest.FoundDomains.'ad.evotec.xyz'.DomainRIDs | Format-Table -AutoSize

$PathToPasswords = 'C:\Users\pklys\OneDrive - Evotec\Support\GitHub\PSWinDocumentation\Ignore\Passwords.txt'
$PathToPasswordsHashes = 'C:\Users\pklys\Downloads\pwned-passwords-ntlm-ordered-by-count\pwned-passwords-ntlm-ordered-by-count.txt'



$Domain = Get-WinADDomainInformation -Domain 'ad.evotec.xyz' -Verbose -PathToPasswords $PathToPasswords #-PathToPasswordsHashes $PathToPasswordsHashes
$Domain.DomainPasswordClearTextPassword | Format-Table -Autosize
$Domain.DomainPasswordLMHash | Format-Table -Autosize
$Domain.DomainPasswordEmptyPassword | Format-Table -Autosize
$Domain.DomainPasswordWeakPassword | Format-Table -Autosize
$Domain.DomainPasswordDefaultComputerPassword | Format-Table -Autosize
$Domain.DomainPasswordPasswordNotRequired | Format-Table -Autosize
$Domain.DomainPasswordPasswordNeverExpires | Format-Table -Autosize
$Domain.DomainPasswordAESKeysMissing | Format-Table -Autosize
$Domain.DomainPasswordPreAuthNotRequired | Format-Table -Autosize
$Domain.DomainPasswordDESEncryptionOnly | Format-Table -Autosize
$Domain.DomainPasswordDelegatableAdmins | Format-Table -Autosize
$Domain.DomainPasswordDuplicatePasswordGroups | Format-Table -Autosize
$Domain.DomainPasswordHashesClearTextPassword | Format-Table -Autosize
$Domain.DomainPasswordHashesLMHash | Format-Table -Autosize
$Domain.DomainPasswordHashesEmptyPassword | Format-Table -Autosize
$Domain.DomainPasswordHashesWeakPassword | Format-Table -Autosize
$Domain.DomainPasswordHashesDefaultComputerPassword | Format-Table -Autosize
$Domain.DomainPasswordHashesPasswordNotRequired | Format-Table -Autosize
$Domain.DomainPasswordHashesPasswordNeverExpires | Format-Table -Autosize
$Domain.DomainPasswordHashesAESKeysMissing | Format-Table -Autosize
$Domain.DomainPasswordHashesPreAuthNotRequired | Format-Table -Autosize
$Domain.DomainPasswordHashesDESEncryptionOnly | Format-Table -Autosize
$Domain.DomainPasswordHashesDelegatableAdmins | Format-Table -Autosize
$Domain.DomainPasswordHashesDuplicatePasswordGroups | Format-Table -Autosize

$Domain.DomainPasswordStats | ft -a
$Domain.DomainPasswordHashesStats | ft -a

#>

#Get-ObjectCount ($Domain.DomainPasswordDuplicatePasswordGroups.'Duplicate Group' | Sort-Object -Unique)

#$Domain.DomainPasswordClearTextPassword | ft -a

<#

$Stats = [ordered] @{}
$Stats.'Clear Text Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordClearTextPassword
$Stats.'LM Hashes' = Get-ObjectCount -Object $Domain.DomainPasswordLMHash
$Stats.'Empty Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordEmptyPassword
$Stats.'Weak Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordWeakPassword
$Stats.'Default Computer Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordDefaultComputerPassword
$Stats.'Password Not Required' = Get-ObjectCount -Object $Domain.DomainPasswordPasswordNotRequired
$Stats.'Password Never Expires' = Get-ObjectCount -Object $Domain.DomainPasswordPasswordNeverExpires
$Stats.'AES Keys Missing' = Get-ObjectCount -Object $Domain.DomainPasswordAESKeysMissing
$Stats.'PreAuth Not Required' = Get-ObjectCount -Object $Domain.DomainPasswordPreAuthNotRequired
$Stats.'DES Encryption Only' = Get-ObjectCount -Object $Domain.DomainPasswordDESEncryptionOnly
$Stats.'Delegatable Admins' = Get-ObjectCount -Object $Domain.DomainPasswordDelegatableAdmins
$Stats.'Duplicate Password Groups' = Get-ObjectCount -Object $Domain.DomainPasswordDuplicatePasswordGroups
$Stats | Ft -a

$StatsHash = [ordered] @{}
$StatsHash.'Clear Text Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordHashesClearTextPassword
$StatsHash.'LM Hashes' = Get-ObjectCount -Object $Domain.DomainPasswordHashesLMHash
$StatsHash.'Empty Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordHashesEmptyPassword
$StatsHash.'Weak Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordHashesWeakPassword
$StatsHash.'Default Computer Passwords' = Get-ObjectCount -Object $Domain.DomainPasswordHashesDefaultComputerPassword
$StatsHash.'Password Not Required' = Get-ObjectCount -Object $Domain.DomainPasswordHashesPasswordNotRequired
$StatsHash.'Password Never Expires' = Get-ObjectCount -Object $Domain.DomainPasswordHashesPasswordNeverExpires
$StatsHash.'AES Keys Missing' = Get-ObjectCount -Object $Domain.DomainPasswordHashesAESKeysMissing
$StatsHash.'PreAuth Not Required' = Get-ObjectCount -Object $Domain.DomainPasswordHashesPreAuthNotRequired
$StatsHash.'DES Encryption Only' = Get-ObjectCount -Object $Domain.DomainPasswordHashesDESEncryptionOnly
$StatsHash.'Delegatable Admins' = Get-ObjectCount -Object $Domain.DomainPasswordHashesDelegatableAdmins
$StatsHash.'Duplicate Password Groups' = Get-ObjectCount -Object $Domain.DomainPasswordHashesDuplicatePasswordGroups
$StatsHash | Ft -a
#>

#Get-ObjectCount   # $Domain.DomainPasswordClearTextPassword