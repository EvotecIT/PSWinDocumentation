function Get-WinADForestUPNSuffixes {
    param(
        $Forest
    )
    @(
        $Forest.RootDomain + ' (Primary / Default UPN)'
        if ($Forest.UPNSuffixes) {
            $Forest.UPNSuffixes
        }
    )       
}