Import-Module PSWriteWord
Import-Module PSWinDocumentation -Force
Import-Module PSWriteColor

$Document = [ordered]@{
    Configuration = [ordered] @{
        Prettify       = @{
            CompanyName = 'Evotec'
        }
        Template       = @{
            Use        = $true
            UseBuiltin = $true


        }
        Options        = @{
            OpenDocument = $true
            OpenExcel    = $true
        }
        DisplayConsole = @{
            ShowTime   = $false
            LogFile    = 'C:\Testing.log'
            TimeFormat = 'yyyy-MM-dd HH:mm:ss'
        }
        Debug          = @{
            Verbose = $false
        }
    }
    DocumentAD    = [ordered] @{
        ExportWord    = $true
        ExportExcel   = $true
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"
        FilePathExcel = ""
        Sections      = [ordered] @{
            SectionForest = [ordered] @{
                SectionForestIntroduction = [ordered] @{

                }
                SectionForestSummary      = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Forest Summary'
                    TocListLevel    = 0
                    TocListItemType = [ListItemType]::Numbered
                    TocHeadingType  = [HeadingType]::Heading1
                    TableData       = [TableData]::ForestSummary
                    TableDesign     = [TableDesign]::ColorfulGridAccent5
                    TableTitleMerge = $true
                    TableTitleText  = "Forest Summary"
                    Text            = "Active Directory at <CompanyName> has a forest name <ForestName>. Following table contains forest summary with important information:"
                }
            }
            SectionDomain = [ordered] @{
                SectionDomainInformation = [ordered] @{

                }
            }
        }

    }
    DocumentAD1   = [ordered] @{

    }
}

Start-Documentation -Document $Document -Verbose