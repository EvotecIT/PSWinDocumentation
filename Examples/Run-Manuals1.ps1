Import-Module PSWriteWord
Import-Module PSWriteExcel
Import-Module PSWinDocumentation #-Force
Import-Module PSWriteColor
Import-Module ActiveDirectory

#[ActiveDirectory].GetEnumValues()

#$Forest =  Get-ADUser -Server $Domain -ResultPageSize 500000 -Filter * -Properties *, "msDS-UserPasswordExpiryTimeComputed" | get-member -force -memberType properties | Select-Object * -ExcludeProperty *Certificate, PropertyNames, *Properties, PropertyCount, Certificates, nTSecurityDescriptor

$Test1 = Get-ADObject -SearchBase (Get-ADRootDSE).SchemaNamingContext -Filter {name -like "User"} -Properties MayContain, SystemMayContain |
    Select-Object @{n = "Attributes"; e = {$_.maycontain + $_.systemmaycontain}} | Select-Object -ExpandProperty Attributes | Sort-Object

$Test2 = Get-ADObject -SearchBase (Get-ADRootDSE).SchemaNamingContext -ldapfilter '(systemFlags:1.2.840.113556.1.4.803:=4)' -Properties systemFlags |
    Select-Object Name | Sort-Object Name


$Test3 = Get-ADObject -SearchBase (Get-ADRootDSE).SchemaNamingContext -Filter *
#$Test3.Properties

$Test4 = $Test3 | Where { $_.Name -like '*User*' }
#$test4

$Domain = 'ad.evotec.xyz'
$Test = Get-ADUser -Server $Domain -ResultPageSize 500000 -Filter * -Properties *, "Exten*", "msDS-UserPasswordExpiryTimeComputed" | Select-Object * -ExcludeProperty *Certificate, PropertyNames, *Properties, PropertyCount, Certificates, nTSecurityDescriptor

$L = Get-ObjectProperties -Object $Test
$L -join ','
$L.Count


#$Test -join ', '

#>
#Format-TransposeTable -Object $Forest.ForestInformation | ft -a
#Format-TransposeTable -Object $Forest.ForestFSMO | ft -AutoSize
#$Forest.FoundDomains
#$Domain = Get-WinADDomainInformation -Domain 'ad.evotec.xyz'
#$Domain

$properties = Get-ADObject -SearchBase (Get-ADRootDSE).SchemanamingContext -Filter {name -eq "User"} -Properties MayContain, SystemMayContain | `
    Select-Object @{name = "Properties"; expression = {$_.maycontain + $_.systemmaycontain}} | Select-Object -ExpandProperty Properties

$Test1 = Get-ADUser -Server $Domain -ResultPageSize 5000000 -Filter * -Properties $properties

$O = Get-ObjectProperties -Object $Test1
$O -join ', '
$O.Count