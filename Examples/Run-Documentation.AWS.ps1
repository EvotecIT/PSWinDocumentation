Import-Module PSWinDocumentation
Import-Module PSWinDocumentation.AD
Import-Module AWSPowerShell
# Import-Module DbaTools # (if you require SQL export)

# Those should automatically load
#Import-Module PSWriteWord
#Import-Module PSWriteExcel

$Document = [ordered]@{
    Configuration = [ordered] @{
        Prettify       = @{
            CompanyName        = 'Evotec'
            UseBuiltinTemplate = $true
            CustomTemplatePath = ''
            Language           = 'en-US'
        }
        Options        = @{
            OpenDocument = $true
            OpenExcel    = $true
        }
        DisplayConsole = @{
            ShowTime   = $false
            LogFile    = "$ENV:TEMP\PSWinDocumentationTesting.log"
            TimeFormat = 'yyyy-MM-dd HH:mm:ss'
        }
        Debug          = @{
            Verbose = $false
        }
    }
    DocumentAWS   = [ordered] @{
        Enable        = $true
        ExportWord    = $true
        ExportExcel   = $true
        ExportSql     = $false
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ReportAWS.docx"
        FilePathExcel = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ReportAWS.xlsx"
        Configuration = [ordered] @{
            AWSAccessKey = ''
            AWSSecretKey = ''
            AWSRegion    = ''
        }
        Sections      = [ordered] @{
            SectionTOC             = [ordered] @{
                Use                  = $true
                TocGlobalDefinition  = $true
                TocGlobalTitle       = 'Table of content'
                TocGlobalRightTabPos = 15
                #TocGlobalSwitches    = 'A', 'C' #[TableContentSwitches]::C, [TableContentSwitches]::A
                PageBreaksAfter      = 1
            }
            SectionAWSIntroduction = [ordered] @{
                ### Enables section
                Use             = $true
                ### Decides how TOC should be visible
                TocEnable       = $True
                TocText         = 'Scope'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                ### Text is added before table/list
                Text            = "This document provides starting overview of AWS..."
                TextAlignment   = [Alignment]::Both
                PageBreaksAfter = 1
            }
            SectionEC2             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - EC2'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [PSWinDocumentation.AWS]::AWSEC2Details
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "Basic information about EC2 servers such as ID, name, environment, instance type and IP."
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSEC2Details'
                ExcelData       = [PSWinDocumentation.AWS]::AWSEC2Details
                SqlExport       = $true
                SqlServer       = 'EVO1'
                SqlDatabase     = 'SSAE18'
                SqlData         = [PSWinDocumentation.AWS]::AWSEC2Details
                SqlTable        = 'dbo.[AWSEC2Details]'
                SqlTableCreate  = $true
            }
            SectionRDS             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - RDS Details'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [PSWinDocumentation.AWS]::AWSRDSDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "Basic information about RDS databases such as name, class, mutliAZ, engine version."
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSRDSDetails'
                ExcelData       = [PSWinDocumentation.AWS]::AWSRDSDetails
            }
            SectionELB             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - Load Balancers'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [PSWinDocumentation.AWS]::AWSLBDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "Basic information about ELB and ALB such as name, DNS name, targets, scheme."
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSLBDetailsList'
                ExcelData       = [PSWinDocumentation.AWS]::AWSLBDetails
            }
            SectionVPC             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - Networking'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [PSWinDocumentation.AWS]::AWSSubnetDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "Basic information about subnets such as: id, name, CIDR, free IP and VPC."
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSSubnetDetails'
                ExcelData       = [PSWinDocumentation.AWS]::AWSSubnetDetails
            }
            SectionEIP             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - Elastic IPs'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [PSWinDocumentation.AWS]::AWSElasticIpDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "Basic information about reserved elastic IPs such as name, IP, network interface."
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSElasticIpDetails'
                ExcelData       = [PSWinDocumentation.AWS]::AWSElasticIpDetails
            }
            SectionIAM             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - IAM Users'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [PSWinDocumentation.AWS]::AWSIAMDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "Basic information about IAM users such as groups and MFA details."
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSIAMDetails'
                ExcelData       = [PSWinDocumentation.AWS]::AWSIAMDetails
            }
        }
    }
}
Start-Documentation -Document $Document -Verbose