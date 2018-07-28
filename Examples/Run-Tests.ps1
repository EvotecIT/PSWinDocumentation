Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord # -Force
Import-Module ActiveDirectory

$Forest = Get-WinADForest
foreach ($Domain in $Forest.Domains) {
    # $ADDomain = Get-ActiveDirectoryCleanData -Domain $Domain
    # Get-PrivilegedGroupsMembers -Domain $ADDomain.DomainInformation.DNSRoot $ADDomain.DomainInformation.DomainSid -Verbose | ft -a
}
$Forest = Get-WinADForestInformation
$Forest.GlobalCatalogs
# $ADSnapshot = Get-ActiveDirectoryCleanData -Domain $Domain
#$AD = Get-ActiveDirectoryProcessedData -ADSnapshot $ADSnapshot

# $ADSnapshot.DomainTrusts

#  $ADSnapshot.DomainControllers | Select Name, Ipv4Address, Ipv6Address, IsGlobalCatalog, IsReadOnly, OperatingSystem, OperationMasterRoles, Site, LdapPort, SSLPort
