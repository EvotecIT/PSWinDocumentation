Import-Module PSWriteWord -Force
Import-Module PSWinDocumentation -Force
Import-Module PSWriteColor

$Document = [ordered]@{
    Configuration    = [ordered] @{
        Prettify       = @{
            CompanyName        = 'Evotec'
            UseBuiltinTemplate = $false
            CustomTemplatePath = ''
            Language           = 'en-US'
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
    DocumentAD       = [ordered] @{
        Enable        = $true
        ExportWord    = $true
        ExportExcel   = $false
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"
        FilePathExcel = ""
        Sections      = [ordered] @{
            SectionForest = [ordered] @{
                SectionTOC                    = [ordered] @{
                    Use                  = $true
                    TocGlobalDefinition  = $true
                    TocGlobalTitle       = 'Table of content'
                    TocGlobalRightTabPos = 15
                    #TocGlobalSwitches    = 'A', 'C' #[TableContentSwitches]::C, [TableContentSwitches]::A
                    PageBreaksAfter      = $true
                }
                SectionForestIntroduction     = [ordered] @{
                    ### Enables section
                    Use             = $true

                    ### Decides how TOC should be visible
                    TocEnable       = $True
                    TocText         = 'Scope'
                    TocListLevel    = 0
                    TocListItemType = [ListItemType]::Numbered
                    TocHeadingType  = [HeadingType]::Heading1

                    ### Text is added before table/list
                    Text            = "This document provides a low-level design of roles and permissions for" `
                        + " the IT infrastructure team at <CompanyName> organization. This document utilizes knowledge from" `
                        + " AD General Concept document that should be delivered with this document. Having all the information" `
                        + " described in attached document one can start designing Active Directory with those principles in mind." `
                        + " It's important to know while best practices that were described are important in decision making they" `
                        + " should not be treated as final and only solution. Most important aspect is to make sure company has full" `
                        + " usability of Active Directory and is happy with how it works. Making things harder just for the sake of" `
                        + " implementation of best practices isn't always the best way to go."
                    TextAlignment   = [Alignment]::Both

                }
                SectionForestSummary          = [ordered] @{
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
                    Text            = "Active Directory at <CompanyName> has a forest name <ForestName>." `
                        + " Following table contains forest summary with important information:"
                }
                SectionForestFSMO             = [ordered] @{
                    Use                   = $true
                    TableData             = [TableData]::ForestFSMO
                    TableDesign           = 'ColorfulGridAccent5'
                    TableTitleMerge       = $true
                    TableTitleText        = 'FSMO Roles'
                    Text                  = 'Following table contains FSMO servers'
                    EmptyParagraphsBefore = 1
                }
                SectionForestOptionalFeatures = [ordered] @{
                    Use                   = $true
                    TableData             = [TableData]::ForestOptionalFeatures
                    TableDesign           = [TableDesign]::ColorfulGridAccent5
                    TableTitleMerge       = $true
                    TableTitleText        = 'Optional Features'
                    Text                  = 'Following table contains optional forest features'
                    EmptyParagraphsBefore = 1
                }
                SectionForestUPNSuffixes      = [ordered] @{
                    Use                   = $true
                    Text                  = "Following UPN suffixes were created in this forest:"
                    ListTextEmpty         = "No UPN suffixes were created in this forest."
                    ListType              = 'Bulleted'
                    ListData              = [TableData]::ForestUPNSuffixes
                    EmptyParagraphsBefore = 1
                }
                SectionForesSPNSuffixes       = [ordered] @{
                    Use                   = $true
                    Text                  = "Following SPN suffixes were created in this forest:"
                    ListTextEmpty         = "No SPN suffixes were created in this forest."
                    ListType              = 'Bulleted'
                    ListData              = [TableData]::ForestSPNSuffixes
                    EmptyParagraphsBefore = 1
                }
                SectionForestSites1           = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Sites'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading1'
                    TableData       = [TableData]::ForestSites1
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = "Forest Sites list can be found below"
                }
                SectionForestSites2           = [ordered] @{
                    Use                   = $true
                    TableData             = [TableData]::ForestSites2
                    TableDesign           = 'ColorfulGridAccent5'
                    Text                  = "Forest Sites list can be found below"
                    EmptyParagraphsBefore = 1
                }
                SectionForestSubnets1         = [ordered] @{
                    Use                   = $true
                    TocEnable             = $True
                    TocText               = 'General Information - Subnets'
                    TocListLevel          = 1
                    TocListItemType       = 'Numbered'
                    TocHeadingType        = 'Heading1'
                    TableData             = [TableData]::ForestSubnets1
                    TableDesign           = 'ColorfulGridAccent5'
                    Text                  = "Table below contains information regarding relation between Subnets and sites"
                    EmptyParagraphsBefore = 1
                }
                SectionForestSubnets2         = [ordered] @{
                    Use                   = $true
                    TableData             = [TableData]::ForestSubnets2
                    TableDesign           = 'ColorfulGridAccent5'
                    Text                  = "Table below contains information regarding relation between Subnets and sites"
                    EmptyParagraphsBefore = 1
                }
                SectionForestSiteLinks        = [ordered] @{
                    Use             = $true
                    TocEnable       = $True
                    TocText         = 'General Information - Site Links'
                    TocListLevel    = 1
                    TocListItemType = 'Numbered'
                    TocHeadingType  = 'Heading1'
                    TableData       = [TableData]::ForestSiteLinks
                    TableDesign     = 'ColorfulGridAccent5'
                    Text            = "Forest Site Links information is available in table below"

                }

            }
            #SectionDomain = [ordered] @{
            #    SectionDomainInformation = [ordered] @{
            #
            #               }
            #          }
        }

    }
    DocumentExchange = [ordered] @{

    }
}

Start-Documentation -Document $Document -Verbose