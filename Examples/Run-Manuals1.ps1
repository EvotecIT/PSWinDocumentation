Import-Module PSWriteWord
Import-Module PSWriteExcel
Import-Module PSWinDocumentation
Import-Module PSWriteColor
Import-Module ActiveDirectory

#$Forest = Get-WinADForestInformation -Verbose
Format-TransposeTable -Object $Forest.ForestInformation | ft -a
Format-TransposeTable -Object $Forest.ForestFSMO | ft -AutoSize
#$Forest.FoundDomains
#$Domain = Get-WinADDomainInformation -Domain 'ad.evotec.xyz'
#$Domain