Import-Module PSWriteWord
Import-Module PSWriteExcel
Import-Module PSWinDocumentation
Import-Module PSWriteColor
Import-Module PSSharedGoods
Import-Module AWSPowerShell

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
        ExportExcel   = $false
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ReportAWS.docx"
        FilePathExcel = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ReportAWS.xlsx"
        Configuration = [ordered] @{
            AWSAccessKey = ''
            AWSSecretKey = ''
            AWSRegion    = ''
        }
        Sections      = [ordered] @{
            SectionTOC              = [ordered] @{
                Use                  = $true
                TocGlobalDefinition  = $true
                TocGlobalTitle       = 'Table of content'
                TocGlobalRightTabPos = 15
                #TocGlobalSwitches    = 'A', 'C' #[TableContentSwitches]::C, [TableContentSwitches]::A
                PageBreaksAfter      = 1
            }
            SectionAWSIntroduction  = [ordered] @{
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
            SectionAWS1             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - AWSEC2Details'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [AWS]::AWSEC2Details
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "AWSEC2Details"
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSEC2Details'
                ExcelData       = [AWS]::AWSEC2Details
            }
            SectionAWS2             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - AWSRDSDetails'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [AWS]::AWSRDSDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5

                Text            = "AWSRDSDetails"
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSRDSDetails'
                ExcelData       = [AWS]::AWSRDSDetails
            }
            SectionAWS3             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - AWSLBDetailsList'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [AWS]::AWSLBDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5

                Text            = "AWSLBDetailsList"
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSLBDetailsList'
                ExcelData       = [AWS]::AWSLBDetails
            }
            SectionAWS4             = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - AWSNetworkDetailsList'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [AWS]::AWSSubnetDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5

                Text            = "AWSSubnetDetails"
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSSubnetDetails'
                ExcelData       = [AWS]::AWSSubnetDetails
            }
            SectionAWSE5            = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - AWSElasticIpDetails'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [AWS]::AWSElasticIpDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "AWSElasticIpDetails"
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSElasticIpDetails'
                ExcelData       = [AWS]::AWSElasticIpDetails
            }
            SectionAWS6DoesntMatter = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - AWSIAMDetails'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [AWS]::AWSIAMDetails
                TableDesign     = [TableDesign]::ColorfulGridAccent5
                Text            = "AWSIAMDetails"
                ExcelExport     = $true
                ExcelWorkSheet  = 'AWSIAMDetails'
                ExcelData       = [AWS]::AWSIAMDetails
            }
        }

    }
}

Start-Documentation -Document $Document -Verbose