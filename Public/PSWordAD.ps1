function Get-DomainPasswordPolicies {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot
    )
    $Paragraph = Add-WordParagraph -WordDocument $WordDocument
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text 'Following table contains password policies'
    $Table = Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $ActiveDirectorySnapshot.DefaultPassWordPoLicy -AutoFit Window -DoNotAddTitle -Design TableGrid
    $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
    $TableParagraph = Add-WordText -WordDocument $WordDocument -Paragraph $TableParagraph -Text 'Default Password Policy' -Alignment center -Color Black -AppendToExistingParagraph
    return $Table
}

function Get-DomainSummary {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot
    )

    $ForestName = $($ActiveDirectorySnapshot.ForestInformation.Name)
    $ForestNameDN = $($ActiveDirectorySnapshot.RootDSE.defaultNamingContext)
    $DomainNetBios = $ActiveDirectorySnapshot.DomainInformation.NetBIOS
    $DomainName = $ActiveDirectorySnapshot.DomainInformation.DNSRoot
    $DomainDistinguishedName = $ActiveDirectorySnapshot.DomainInformation.DistinguishedName

    $Text = "Following domains exists within forest $ForestName"
    $Text0 = "Domain $DomainDistinguishedName"
    $Text1 = "Name for fully qualified domain name (FQDN): $DomainName"
    $Text2 = "Name for NetBIOS: $DomainNetBios"

    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text

    $ListDomainInformation = $null
    $ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel 0 -ListItemType Bulleted -ListValue $Text0
    $ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel 1 -ListItemType Bulleted -ListValue $Text1
    $ListDomainInformation = $ListDomainInformation | New-WordListItem -WordDocument $WordDocument -ListLevel 1 -ListItemType Bulleted -ListValue $Text2
    Add-WordListItem -WordDocument $WordDocument -Paragraph $Paragraph -List $ListDomainInformation -Verbose -Supress $true
}

function Get-DomainUPNSuffixes {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot
    )
    $Text = "Following UPN suffixes were created for $CompanyName users and used as part of user's logon processes:"
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text
    $List = Add-WordList -WordDocument $WordDocument -ListType Bulleted -Paragraph $Paragraph -ListData $ActiveDirectorySnapshot.UPNSuffixes #-Verbose
    return $List
}

function Get-ForestFeatures {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot
    )

    $Paragraph = Add-WordParagraph -WordDocument $WordDocument
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text 'Following table contains Forest Features'
    $Table = Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $ActiveDirectorySnapshot.OptionalFeatures -AutoFit Window -DoNotAddTitle -Design TableGrid
    $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
    $TableParagraph = Add-WordText -WordDocument $WordDocument -Paragraph $TableParagraph -Text 'Forest Features' -Alignment center -Color Black -AppendToExistingParagraph
    return $Table
}

function Get-ForestFSMORoles {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot
    )
    $Paragraph = Add-WordParagraph -WordDocument $WordDocument
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text 'Following table contains FSMO servers'
    $Table = Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $ActiveDirectorySnapshot.FSMO -AutoFit Window -DoNotAddTitle -Design TableGrid
    $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
    $TableParagraph = Add-WordText -WordDocument $WordDocument -Paragraph $TableParagraph -Text 'FSMO Roles' -Alignment center -Color Black -AppendToExistingParagraph

    return $Table
}

function Get-ForestSummary {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot
    )

    $ForestName = $($ActiveDirectorySnapshot.ForestInformation.Name)
    $ForestNameDN = $($ActiveDirectorySnapshot.RootDSE.defaultNamingContext)

    $ForestSummaryText = "Active Directory at $CompanyName has a forest name $ForestName ($ForestNameDN). Following table contains forest summary with important information:"
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $ForestSummaryText
    $Table = Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $ActiveDirectorySnapshot.ForestInformation -AutoFit Window -DoNotAddTitle  -Design TableGrid
    $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
    $TableParagraph = Add-WordText -WordDocument $WordDocument -Paragraph $TableParagraph -Text 'Forest Summary' -Alignment center -Color Black -Append
    return $Table
}

function Get-DocumentScope {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot,
        $CompanyName
    )

    $Text = "This document provides a low-level design of roles and permissions for the IT infrastructure team at $CompanyName organization. This document utilizes knowledge from AD General Concept document that should be delivered with this document. Having all the information described in attached document one can start designing Active Directory with those principles in mind. It's important to know while best practices that were described are important in decision making they should not be treated as final and only solution. Most important aspect is to make sure company has full usability of Active Directory and is happy with how it works. Making things harder just for the sake of implementation of best practices isn't always the best way to go."
    #$Section1 = Add-WordParagraph -WordDocument $WordDocument -Paragraph $Section1 -AfterSelf -Supress $false
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text -Alignment both
    return $Paragraph
}
