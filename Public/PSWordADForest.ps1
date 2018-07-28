
function Get-ForestUPNSuffixes {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $CompanyName,
        $ActiveDirectorySnapshot,
        $Domain
    )
    $Text = "Following UPN suffixes were created for $CompanyName users and used as part of user's logon processes:"
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text
    $List = Add-WordList -WordDocument $WordDocument -ListType Bulleted -Paragraph $Paragraph -ListData $ActiveDirectorySnapshot.UPNSuffixes #-Verbose
    return $List
}

function Get-ForestSPNSuffixes {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $CompanyName,
        $ActiveDirectorySnapshot,
        $Domain
    )
    if ((Get-ObjectCount $ActiveDirectorySnapshot.SPNSuffixes) -gt 0) {
        $Text = "Following SPN suffixes were created in this forest:"
        $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text
        $List = Add-WordList -WordDocument $WordDocument -ListType Bulleted -Paragraph $Paragraph -ListData $ActiveDirectorySnapshot.SPNSuffixes #-Verbose
    } else {
        $Text = "No SPN suffixes were created in this forest."
        $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $Text
    }
    return $List
}

function Get-ForestSummary {
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)]$Paragraph,
        $ActiveDirectorySnapshot
    )

    $ForestName = $($ActiveDirectorySnapshot.ForestInformation.Name)
    $ForestNameDN = $($ActiveDirectorySnapshot.RootDSE.defaultNamingContext)

    $ForestSummaryText = "Active Directory at $CompanyName has a forest name $ForestName. Following table contains forest summary with important information:"
    $Paragraph = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text $ForestSummaryText
    $Table = Add-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -DataTable $ActiveDirectorySnapshot.ForestInformation -AutoFit Window -DoNotAddTitle  -Design ColorfulGridAccent5
    $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -ColumnNrStart 0 -ColumnNrEnd 1
    $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
    $TableParagraph = Add-WordText -WordDocument $WordDocument -Paragraph $TableParagraph -Text 'Forest Summary' -Alignment center -Color Black -Append
    return $Table
}