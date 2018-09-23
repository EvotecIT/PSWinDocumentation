Import-Module PSWriteWord
Import-Module PSWriteExcel
Import-Module PSWinDocumentation
Import-Module PSWriteColor
Import-Module PSSharedGoods
Import-Module AWSPowerShell

$Document = [ordered]@{
    Configuration     = [ordered] @{
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
    DocumentOffice365 = [ordered] @{
        Enable        = $true
        ExportWord    = $true
        ExportExcel   = $true
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ReportO365.docx"
        FilePathExcel = "$Env:USERPROFILE\Desktop\PSWinDocumentation-ReportO365.xlsx"
        Configuration = [ordered] @{
            O365ExchangeUse            = $true
            O365AzureADUse             = $true

            O365Username               = 'przemyslaw.klys@evotec.pl'
            O365Password               = 'C:\Users\pklys\OneDrive - Evotec\Support\GitHub\PSWinDocumentation\Ignore\MySecurePassword.txt'
            O365PasswordAsSecure       = $true
            O365PasswordFromFile       = $true

            O365ExchangeSessionName    = 'O365ExchangeOnline'
            O365ExchangeAuthentication = 'Basic'
            O365ExchangeURI            = 'https://outlook.office365.com/powershell-liveid/'

            O365AzureSessionName       = 'O365Azure'

        }
        Sections      = [ordered] @{
            SectionO365TOC                          = [ordered] @{
                Use                  = $true
                TocGlobalDefinition  = $true
                TocGlobalTitle       = 'Table of content'
                TocGlobalRightTabPos = 15
                #TocGlobalSwitches    = 'A', 'C' #[TableContentSwitches]::C, [TableContentSwitches]::A
                PageBreaksAfter      = 0
            }
            SectionO365Introduction                 = [ordered] @{
                ### Enables section
                Use             = $true

                ### Decides how TOC should be visible
                TocEnable       = $True
                TocText         = 'Scope'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1

                ### Text is added before table/list
                Text            = ""
                TextAlignment   = [Alignment]::Both
                PageBreaksAfter = 1

            }
            SectionO365ExchangeMailBoxes            = [ordered] @{
                Use             = $true
                TocEnable       = $True
                TocText         = 'General Information - O365ExchangeMailBoxes'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1
                TableData       = [O365]::O365UExchangeMailBoxes
                TableDesign     = [TableDesign]::ColorfulGridAccent5

                Text            = "This is test"
                ExcelExport     = $false
                ExcelWorkSheet  = 'O365ExchangeMailBoxes'
                ExcelData       = [O365]::O365UExchangeMailBoxes
            }
            O365AzureTenantDomains                  = [ordered] @{
                Use                 = $true

                TocEnable           = $True
                TocText             = 'General Information - Office 365 Domains'
                TocListLevel        = 0
                TocListItemType     = [ListItemType]::Numbered
                TocHeadingType      = [HeadingType]::Heading1

                TableData           = [O365]::O365AzureTenantDomains
                TableMaximumColumns = 7
                TableDesign         = [TableDesign]::ColorfulGridAccent5

                Text                = 'Following table contains all domains defined in Office 365 portal.'

                ExcelExport         = $true
                ExcelWorkSheet      = 'O365 Domains'
                ExcelData           = [O365]::O365AzureTenantDomains
            }
            O365AzureADGroupMembersUser             = [ordered] @{
                Use             = $true

                TocEnable       = $True
                TocText         = 'General Information - O365AzureADGroupMembersUser'
                TocListLevel    = 0
                TocListItemType = [ListItemType]::Numbered
                TocHeadingType  = [HeadingType]::Heading1

                TableData       = [O365]::O365AzureADGroupMembersUser
                TableDesign     = [TableDesign]::ColorfulGridAccent5

                ExcelExport     = $true
                ExcelWorkSheet  = 'O365AzureADGroupMembersUser'
                ExcelData       = [O365]::O365AzureADGroupMembersUser
            }


            ## Data below makes sense only in Excel / SQL Export
            O365AzureLicensing                      = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365AzureLicensing'
                ExcelData      = [O365]::O365UAzureLicensing
            }
            O365AzureSubscription                   = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365AzureSubscription'
                ExcelData      = [O365]::O365UAzureSubscription
            }
            O365AzureADUsers                        = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365AzureADUsers'
                ExcelData      = [O365]::O365UAzureADUsers
            }
            O365AzureADUsersDeleted                 = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365AzureADUsersDeleted'
                ExcelData      = [O365]::O365UAzureADUsersDeleted
            }
            O365AzureADGroups                       = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365AzureADGroups'
                ExcelData      = [O365]::O365UAzureADGroups
            }
            O365AzureADGroupMembers                 = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365AzureADGroupMembers'
                ExcelData      = [O365]::O365UAzureADGroupMembers
            }
            O365AzureADContacts                     = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365AzureADContacts'
                ExcelData      = [O365]::O365UAzureADContacts
            }


            O365ExchangeMailBoxes                   = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeMailBoxes'
                ExcelData      = [O365]::O365UExchangeMailBoxes
            }
            O365ExchangeMailUsers                   = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeMailUsers'
                ExcelData      = [O365]::O365UExchangeMailUsers
            }
            O365ExchangeRecipientsPermissions       = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeRecipientsPermissions'
                ExcelData      = [O365]::O365UExchangeRecipientsPermissions
            }
            O365ExchangeGroupsDistributionDynamic   = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeGroupsDistributionDynamic'
                ExcelData      = [O365]::O365UExchangeGroupsDistributionDynamic
            }
            O365ExchangeMailboxesEquipment          = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeMailboxesEquipment'
                ExcelData      = [O365]::O365UExchangeMailboxesEquipment
            }
            O365ExchangeUsers                       = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeUsers'
                ExcelData      = [O365]::O365UExchangeUsers
            }
            O365ExchangeMailboxesRooms              = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeMailboxesRooms'
                ExcelData      = [O365]::O365UExchangeMailboxesRooms
            }
            O365ExchangeGroupsDistributionMembers   = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeGroupsDistributionMembers'
                ExcelData      = [O365]::O365UExchangeGroupsDistributionMembers
            }
            O365ExchangeEquipmentCalendarProcessing = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeEquipmentCalendarProcessing'
                ExcelData      = [O365]::O365UExchangeEquipmentCalendarProcessing
            }
            O365ExchangeGroupsDistribution          = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeGroupsDistribution'
                ExcelData      = [O365]::O365UExchangeGroupsDistribution
            }
            O365ExchangeContactsMail                = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeContactsMail'
                ExcelData      = [O365]::O365UExchangeContactsMail
            }
            O365ExchangeMailboxesJunk               = [ordered] @{
                Use            = $false
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeMailboxesJunk'
                ExcelData      = [O365]::O365UExchangeMailboxesJunk
            }
            O365ExchangeRoomsCalendarPrcessing      = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeRoomsCalendarPrcessing'
                ExcelData      = [O365]::O365UExchangeRoomsCalendarPrcessing
            }
            O365ExchangeContacts                    = [ordered] @{
                Use            = $true
                ExcelExport    = $true
                ExcelWorkSheet = 'O365ExchangeContacts'
                ExcelData      = [O365]::O365UExchangeContacts
            }
        }
    }
}

Start-Documentation -Document $Document -Verbose