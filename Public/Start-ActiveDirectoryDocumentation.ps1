function Start-ActiveDirectoryDocumentation {
    [CmdletBinding()]
    param (
        [string] $FilePath,
        [string] $FilePathExcel,
        [switch] $CleanDocument,
        [string] $CompanyName = 'Evotec',
        [switch] $OpenDocument,
        [switch] $OpenExcel
    )
    Write-Warning 'This is legacy command. Use Start-Documentation instead. Will be terminated sooner or later'
    # Left here for legacy reasons.
    $Document = $Script:Document
    $Document.Configuration.Prettify.CompanyName = $CompanyName
    if ($CleanDocument) {
        $Document.Configuration.Prettify.UseBuiltinTemplate = $false
    }
    $Document.Configuration.Options.OpenDocument = $OpenDocument
    $Document.Configuration.Options.OpenExcel = $OpenExcel
    $Document.DocumentAD.FilePathWord = $FilePath
    if ($FilePathExcel) {
        $Document.DocumentAD.ExportExcel = $true
        $Document.DocumentAD.FilePathExcel = $FilePathExcel
    } else {
        $Document.DocumentAD.ExportExcel = $false
    }

    Start-Documentation -Document $Document
}