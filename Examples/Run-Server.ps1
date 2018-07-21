Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord -Force
Import-Module ActiveDirectory

$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate.docx"

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"

function Start-WinDocumentationServer {
    param (
        [string[]] $ComputerName = $Env:COMPUTERNAME,
        [string] $FilePathTemplate,
        [string] $FilePath,
        [switch] $OpenDocument
    )
    $WordDocument = New-WordDocument -FilePath $FilePath
    $ListOfHeaders = @(
        'Scope',
        'General Information',
        'General Information - Forest Summary',
        'General Information - Domain Summary')

    $Toc = Add-WordToc -WordDocument $WordDocument -Title 'Table of content' -Switches C, A -RightTabPos 15 -HeaderStyle Heading1

    ### This list will be converted into Headings for Numbered Table of Contents
    $ListHeaders = Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfHeaders
    ### Converts List into numbered Headings for Table of Content
    $Headings = Convert-ListToHeadings -WordDocument $WordDocument -List $ListHeaders
    $Text = "This document provides a low-level design of roles and permissions for the IT infrastructure team at Colmore for a 3 level support organization. This document utilizes knowledge from AD General Concept document that should be delivered with this document. Having all the information described in attached document one can start designing Active Directory with those principles in mind. It's important to know while best practices that were described are important in decision making they should not be treated as final and only solution. Most important aspect is to make sure company has full usability of Active Directory and is happy with how it works. Making things harder just for the sake of implementation of best practices isn't always the best way to go."
    $Section1 = $Headings[0]
    #$Section1 = Add-WordParagraph -WordDocument $WordDocument -Paragraph $Section1 -AfterSelf -Supress $false
    $Section1 = Add-WordText -WordDocument $WordDocument -Paragraph $Section1 -Text $Text -Alignment both

    ### 3rd section - Forest Summary
    $ActiveDirectorySnapshot = Get-ActiveDirectoryData
    $ActiveDirectoryForestInformation = $ActiveDirectorySnapshot.ForestInformation
    $ActiveDirectoryForestFeatures = $ActiveDirectorySnapshot.ForestFeatures
    $ActiveDirectoryFSMO = $ActiveDirectorySnapshot.FSMO

    $Section3Paragraph1 = $Headings[2]
    $Section3Paragraph1 = Add-WordText -WordDocument $WordDocument -Paragraph $Section3Paragraph1 -Text 'Following table contains Forest Summary'
    $Section3Table1 = Add-WordTable -WordDocument $WordDocument -Paragraph $Section3Paragraph1 -DataTable $ActiveDirectoryForestInformation -AutoFit Window -DoNotAddTitle  -Design TableGrid
    $Section3Table1 = Set-WordTableRowMergeCells -Table $Section3Table1 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $Section3Table1Paragraph = Get-WordTableRow -Table $Section3Table1 -RowNr 0 -ColumnNr 0
    $Section3Table1Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Section3Table1Paragraph -Text 'Forest Summary' -Alignment center -Color Black -Append

    $Section3Paragraph2 = Add-WordParagraph -WordDocument $WordDocument -Table $Section3Table1 -AfterSelf
    $Section3Paragraph2 = Add-WordText -WordDocument $WordDocument -Paragraph $Section3Paragraph2 -Text 'Following table contains Forest Features'
    $Section3Table2 = Add-WordTable -WordDocument $WordDocument -Paragraph $Section3Paragraph2 -DataTable $ActiveDirectoryForestFeatures -AutoFit Window -DoNotAddTitle -Design TableGrid
    $Section3Table2 = Set-WordTableRowMergeCells -Table $Section3Table2 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $Section3Table2Paragraph = Get-WordTableRow -Table $Section3Table2 -RowNr 0 -ColumnNr 0
    $Section3Table2Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Section3Table2Paragraph -Text 'Forest Summary' -Alignment center -Color Black -Append

    $Section3Paragraph3 = Add-WordParagraph -WordDocument $WordDocument -Table $Section3Table2 -AfterSelf
    $Section3Paragraph3 = Add-WordText -WordDocument $WordDocument -Paragraph $Section3Paragraph3 -Text 'Following table contains FSMO servers'
    $Section3Table3 = Add-WordTable -WordDocument $WordDocument -Paragraph $Section3Paragraph3 -DataTable $ActiveDirectoryFSMO -AutoFit Window -DoNotAddTitle -Design TableGrid
    $Section3Table3 = Set-WordTableRowMergeCells -Table $Section3Table3 -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $Section3Table2Paragraph = Get-WordTableRow -Table $Section3Table3 -RowNr 0 -ColumnNr 0
    $Section3Table3Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Section3Table2Paragraph -Text 'FSMO Roles' -Alignment center -Color Black -Append



    <# Implment
    $ColorShading1 = Get-ColorFromARGB 255 0 112 192
    $ColorFill1 = Get-ColorFromARGB 0 0 112 192

    $ColorShading2 = Get-ColorFromARGB 255 231 230 230
    $ColorFill2 = Get-ColorFromARGB 0 231 230 230


    Set-WordTableCell -Table $Section3Table -RowNr 0 -ColumnNr 0 -FillColor $ColorFill1 -ShadingColor $ColorShading1
    Set-WordTableCell -Table $Section3Table -RowNr 1 -ColumnNr 0 -FillColor $ColorFill2 -ShadingColor $ColorShading2
    Set-WordTableCell -Table $Section3Table -RowNr 1 -ColumnNr 1 -FillColor $ColorFill2 -ShadingColor $ColorShading2
    Set-WordTableCell -Table $Section3Table -RowNr 2 -ColumnNr 0 -FillColor $ColorFill2 -ShadingColor $ColorShading2
    Set-WordTableCell -Table $Section3Table -RowNr 2 -ColumnNr 1 -FillColor $ColorFill2 -ShadingColor $ColorShading2

    $BorderTypeTop = New-WordTableBorder -BorderStyle Tcbs_none -BorderSize one -BorderSpace 1 -BorderColor Blue
    $BorderTypeBottom = New-WordTableBorder -BorderStyle Tcbs_none -BorderSize one -BorderSpace 1 -BorderColor Red
    $BorderTypeLeft = New-WordTableBorder -BorderStyle Tcbs_none -BorderSize one -BorderSpace 1 -BorderColor Blue
    $BorderTypeRight = New-WordTableBorder -BorderStyle Tcbs_none -BorderSize one -BorderSpace 1 -BorderColor Yellow
    $BorderTypeInsideH = New-WordTableBorder -BorderStyle Tcbs_none -BorderSize one -BorderSpace 1 -BorderColor Pink
    $BorderTypeInsideV = New-WordTableBorder -BorderStyle Tcbs_none -BorderSize one -BorderSpace 1 -BorderColor Black

    Set-WordTableBorder -Table $Section3Table -TableBorderType Top -Border $BorderTypeTop
    Set-WordTableBorder -Table $Section3Table -TableBorderType Bottom -Border $BorderTypeBottom
    Set-WordTableBorder -Table $Section3Table -TableBorderType Left -Border $BorderTypeLeft
    Set-WordTableBorder -Table $Section3Table -TableBorderType Right -Border $BorderTypeRight
    Set-WordTableBorder -Table $Section3Table -TableBorderType InsideH -Border $BorderTypeInsideH
    Set-WordTableBorder -Table $Section3Table -TableBorderType InsideV -Border $BorderTypeInsideV
    #>


    #$Section3Table.Rows[0].Cells[0].FillColor = "Grey"
    #$Section3Table.Rows[1].Cells[0].FillColor = "Blue"

    #$Section3Table
    #.Text = 'Forest Information'
    #$Section3.Rows[0].Text = 'Forest Information'






    Save-WordDocument -WordDocument $WordDocument -Language 'en-US' -FilePath $FilePath -Supress $true
    if ($OpenDocument) { Invoke-Item $FilePath }
}

Start-WinDocumentationServer -ComputerName 'AD1' -FilePathTemplate $FilePathTemplate -FilePath $FilePath #-OpenDocument