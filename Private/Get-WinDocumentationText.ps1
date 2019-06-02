function Get-WinDocumentationText {
    [CmdletBinding()]
    param (
        [string[]] $Text,
        [hashtable] $Forest,
        [string] $Domain
    )

    $Array = foreach ($T in $Text) {
        $T = $T.Replace('<CompanyName>', $Document.Configuration.Prettify.CompanyName)
        $T = $T.Replace('<ForestName>', $Forest.ForestInformation.Name)
        $T = $T.Replace('<ForestNameDN>', $Forest.ForestInformation.'Forest Distingushed Name')
        $T = $T.Replace('<Domain>', $Domain)
        $T = $T.Replace('<DomainNetBios>', $Forest.FoundDomains.$Domain.DomainInformation.NetBIOSName)
        $T = $T.Replace('<DomainDN>', $Forest.FoundDomains.$Domain.DomainInformation.DistinguishedName)
        $T = $T.Replace('<DomainPasswordWeakPasswordList>', $Forest.FoundDomains.$Domain.DomainPasswordDataPasswords.DomainPasswordWeakPasswordList)
        $T
    }
    return $Array
}