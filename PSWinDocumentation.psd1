@{
    Copyright = 'Evotec (c) 2011-2019. All rights reserved.'
    PrivateData = @{
        PSData = @{
            Tags = 'documentation', 'windows', 'word', 'workstation', 'activedirectory', 'ad', 'excel', 'sql', 'azure', 'azuread', 'exchange', 'office365', 'aws'
            ProjectUri = 'https://github.com/EvotecIT/PSWinDocumentation'
            IconUri = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWinDocumentation.png'
            Prerelease = 'Preview7'
        }
    }
    Description = 'Simple project generating Active Directory, AWS, Exchange, Office 365 (Exchange, Azure AD) documentation to Microsoft Word, Microsoft Excel and Microsoft SQL. More things to follow...'
    PowerShellVersion = '5.1'
    FunctionsToExport = 'Start-Documentation'
    Author = 'Przemyslaw Klys'
    RequiredModules = @{
        ModuleVersion = '0.7.2'
        ModuleName = 'PSWriteWord'
        GUID = '6314c78a-d011-4489-b462-91b05ec6a5c4'
    }, @{
        ModuleVersion = '0.1'
        ModuleName = 'PSWriteExcel'
        GUID = '82232c6a-27f1-435d-a496-929f7221334b'
    }, @{
        ModuleVersion = '0.0.77'
        ModuleName = 'PSSharedGoods'
        GUID = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
    }, @{
        ModuleVersion = '0.0.7'
        ModuleName = 'PSWinDocumentation.AD'
        GUID = 'a46f9775-04d2-4423-9631-01cfda42b95d'
    }, @{
        ModuleVersion = '0.0.3'
        ModuleName = 'PSWinDocumentation.AWS'
        GUID = 'b3c23202-740d-4f7b-a9d7-bd87063381cc'
    }, @{
        ModuleVersion = '0.0.1'
        ModuleName = 'PSWinDocumentation.O365'
        GUID = '71ea1419-d950-444b-83c9-c579de74962a'
    }
    GUID = '6bd80c20-e606-4e31-9f88-9ad305256f23'
    RootModule = 'PSWinDocumentation.psm1'
    AliasesToExport = ''
    ModuleVersion = '0.5.0'
    CompanyName = 'Evotec'
}