Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord # -Force
Import-Module ActiveDirectory

$Forest = Get-WinADForestInformation
foreach ($Domain in $Forest.Domains) {
    # $ADDomain = Get-ActiveDirectoryCleanData -Domain $Domain
    # Get-PrivilegedGroupsMembers -Domain $ADDomain.DomainInformation.DNSRoot $ADDomain.DomainInformation.DomainSid -Verbose | ft -a
}
#$Forest.GlobalCatalogs

$Domain = 'ad.evotec.xyz'
$DomainInformation = Get-WinDomainInformation  -Domain $Domain
#$DomainInformation.ADSnapshot.DomainAdministrators
$DomainInformation.GroupPolicies
#$ADSnapshot.OrganizationalUnits | ft -a
# $ADSnapshot = Get-ActiveDirectoryCleanData -Domain $Domain
#$AD = Get-ActiveDirectoryProcessedData -ADSnapshot $ADSnapshot

# $ADSnapshot.DomainTrusts

#  $ADSnapshot.DomainControllers | Select Name, Ipv4Address, Ipv6Address, IsGlobalCatalog, IsReadOnly, OperatingSystem, OperationMasterRoles, Site, LdapPort, SSLPort
#get-ADOrganizationalUnit -Properties * -Filter * | Select Name, CanonicalName, Created | Sort CanonicalName

#get-process | Where { $_.MainWindowTitle -ne '' } | Select-Object id, name, mainwindowtitle | Sort-Object mainwindowtitle  | Ft -a