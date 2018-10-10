#Clear-Host
Import-Module PSWinDocumentation #-Force
Import-Module PSSharedGoods
Import-Module ActiveDirectory

$PathToPasswords = 'C:\Users\pklys\OneDrive - Evotec\Support\GitHub\PSWinDocumentation\Ignore\Passwords.txt'
$PathToPasswordsHashes = 'C:\Users\pklys\Downloads\pwned-passwords-ntlm-ordered-by-count\pwned-passwords-ntlm-ordered-by-count.txt'

$Forest = Get-WinADForestInformation -Verbose -PathToPasswords $PathToPasswords
#$Forest.FoundDomains.'ad.evotec.pl'
#$Forest.FoundDomains.'ad.evotec.pl'.DomainFineGrainedPoliciesUsers | Format-Table -AutoSize
#$Forest.FoundDomains.'ad.evotec.xyz'.DomainRIDs | Format-Table -AutoSize

# Alternatively ask for domain directly

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
$Domain.DomainPasswordHashesWeakPassword | Format-Table -Autosize

$Domain.DomainPasswordStats | ft -a
$Domain.DomainPasswordHashesStats | ft -a
