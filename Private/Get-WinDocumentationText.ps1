function Get-WinDocumentationText {
    param (
        [string[]] $Text,
        [hashtable] $Forest,
        [string] $Domain
    )
    $Array = @()
    foreach ($T in $Text) {
        $T = $T.Replace('<CompanyName>', $Document.Configuration.Prettify.CompanyName)
        $T = $T.Replace('<ForestName>', $Forest.ForestName)
        $T = $T.Replace('<ForestNameDN>', $Forest.RootDSE.defaultNamingContext)
        $T = $T.Replace('<Domain>', $Domain)
        $T = $T.Replace('<DomainNetBios>', $Forest.FoundDomains.$Domain.DomainInformation.NetBIOSName)
        $T = $T.Replace('<DomainDN>', $Forest.FoundDomains.$Domain.DomainInformation.DistinguishedName)
        $Array += $T
    }
    return $Array
}