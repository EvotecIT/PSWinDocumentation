function Get-DocumentPath {
    [CmdletBinding()]
    param (
        [System.Object] $Document,
        [string] $FinalDocumentLocation
    )
    if ($Document.Configuration.Prettify.UseBuiltinTemplate) {
        #Write-Verbose 'Get-DocumentPath - Option 1'
        #$WordDocument = Get-WordDocument -FilePath "$((get-item $PSScriptRoot).Parent.FullName)\Templates\WordTemplate.docx"
        $WordDocument = Get-WordDocument -FilePath "$($MyInvocation.MyCommand.Module.ModuleBase)\Templates\WordTemplate.docx"
    } else {
        if ($Document.Configuration.Prettify.CustomTemplatePath) {
            if ($(Test-File -File $Document.Configuration.Prettify.CustomTemplatePath -FileName 'CustomTemplatePath') -eq 0) {
                # Write-Verbose 'Get-DocumentPath - Option 2'
                $WordDocument = Get-WordDocument -FilePath $Document.Configuration.Prettify.CustomTemplatePath
            } else {
                #Write-Verbose 'Get-DocumentPath - Option 3'
                $WordDocument = New-WordDocument -FilePath $FinalDocumentLocation
            }
        } else {
            #Write-Verbose 'Get-DocumentPath - Option 4'
            $WordDocument = New-WordDocument -FilePath $FinalDocumentLocation
        }
    }
    if ($WordDocument -eq $null) { Write-Verbose ' Null'}
    return $WordDocument
}
