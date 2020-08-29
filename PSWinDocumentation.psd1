@{
    AliasesToExport   = ''
    Author            = 'Przemyslaw Klys'
    CompanyName       = 'Evotec'
    Copyright         = '(c) 2011 - 2020 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description       = 'Simple project generating Active Directory, AWS, Office 365 (Exchange, Azure AD) documentation to Microsoft Word, Microsoft Excel and Microsoft SQL. More things to follow...'
    FunctionsToExport = 'Invoke-ADExcel', 'Invoke-ADHTML', 'Invoke-ADWord', 'Invoke-Documentation', 'Show-GroupMember', 'Start-Documentation'
    GUID              = '6bd80c20-e606-4e31-9f88-9ad305256f23'
    ModuleVersion     = '0.5.4'
    PowerShellVersion = '5.1'
    PrivateData       = @{
        PSData = @{
            Tags                       = 'documentation', 'windows', 'word', 'workstation', 'activedirectory', 'ad', 'excel', 'sql', 'azure', 'azuread', 'exchange', 'office365', 'aws'
            ProjectUri                 = 'https://github.com/EvotecIT/PSWinDocumentation'
            ExternalModuleDependencies = 'ActiveDirectory', 'GroupPolicy'
            IconUri                    = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWinDocumentation.png'
        }
    }
    RequiredModules   = @{
        ModuleVersion = '0.0.97'
        ModuleName    = 'PSWriteHTML'
        Guid          = 'a7bdf640-f5cb-4acf-9de0-365b322d245c'
    }, @{
        ModuleVersion = '1.1.8'
        ModuleName    = 'PSWriteWord'
        Guid          = '6314c78a-d011-4489-b462-91b05ec6a5c4'
    }, @{
        ModuleVersion = '0.1.10'
        ModuleName    = 'PSWriteExcel'
        Guid          = '82232c6a-27f1-435d-a496-929f7221334b'
    }, @{
        ModuleVersion = '0.0.169'
        ModuleName    = 'PSSharedGoods'
        Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
    }, @{
        ModuleVersion = '0.1.19'
        ModuleName    = 'PSWinDocumentation.AD'
        Guid          = 'a46f9775-04d2-4423-9631-01cfda42b95d'
    }, @{
        ModuleVersion = '0.0.4'
        ModuleName    = 'PSWinDocumentation.AWS'
        Guid          = 'b3c23202-740d-4f7b-a9d7-bd87063381cc'
    }, @{
        ModuleVersion = '0.0.7'
        ModuleName    = 'PSWinDocumentation.O365'
        Guid          = '71ea1419-d950-444b-83c9-c579de74962a'
    }, @{
        ModuleVersion = '0.0.70'
        ModuleName    = 'ADEssentials'
        Guid          = '9fc9fd61-7f11-4f4b-a527-084086f1905f'
    }, @{
        ModuleVersion = '0.0.60'
        ModuleName    = 'GPOZaurr'
        Guid          = 'f7d4c9e4-0298-4f51-ad77-e8e3febebbde'
    }, 'ActiveDirectory', 'GroupPolicy'
    RootModule        = 'PSWinDocumentation.psm1'
}