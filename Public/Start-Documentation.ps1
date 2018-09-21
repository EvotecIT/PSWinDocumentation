function Start-Documentation {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    $TimeTotal = [System.Diagnostics.Stopwatch]::StartNew() # Timer Start
    Test-Configuration -Document $Document

    if ($Document.DocumentAD.Enable) {
        Start-DocumentationAD -Document $Document
    }
    if ($Document.DocumentAWS.Enable) {
        Start-DocumentationAWS -Document $Document
    }
    if ($Document.DocumentExchange.Enable) {
        Start-DocumentationExchange -Document $Document
    }
}