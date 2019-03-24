function Get-WinADForestInformation {
    [CmdletBinding()]
    param (
        [Object] $TypesRequired,
        [switch] $RequireTypes,
        [string] $PathToPasswords,
        [string] $PathToPasswordsHashes
    )
    $TimeToGenerate = Start-TimeLog
    if ($null -eq $TypesRequired) {
        # Gets all types
        Write-Verbose 'Get-WinADForestInformation - TypesRequired is null. Getting all.'
        $TypesRequired = Get-Types -Types ([ActiveDirectory])
    }

    $Data = [ordered] @{ }
    Write-Verbose 'Getting forest information - Forest'
    $Data.Forest = Get-WinForest
    Write-Verbose 'Getting forest information - RootDSE'
    $Data.RootDSE = Get-WinADRootDSE
    Write-Verbose 'Getting forest information - ForestName/ForestNameDN'
    $Data.ForestName = $Data.Forest.Name
    $Data.ForestNameDN = $Data.RootDSE.defaultNamingContext
    $Data.Domains = $Data.Forest.Domains

    ## Forest Information
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestInformation)) {
        Write-Verbose 'Getting forest information - Forest Information'
        $Data.ForestInformation = Get-WinADForestInfo -Forest $Data.Forest
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestUPNSuffixes)) {
        Write-Verbose 'Getting forest information - Forest UPNSuffixes'
        $Data.ForestUPNSuffixes = Get-WinADForestUPNSuffixes -Forest $Data.Forest
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestGlobalCatalogs)) {
        Write-Verbose 'Getting forest information - Forest GlobalCatalogs'
        $Data.ForestGlobalCatalogs = $Data.Forest.GlobalCatalogs
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSPNSuffixes)) {
        Write-Verbose 'Getting forest information - Forest SPNSuffixes'
        $Data.ForestSPNSuffixes = $Data.Forest.SPNSuffixes
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestFSMO)) {
        Write-Verbose 'Getting forest information - Forest FSMO'
        $Data.ForestFSMO = [ordered] @{
            'Domain Naming Master' = $Data.Forest.DomainNamingMaster
            'Schema Master'        = $Data.Forest.SchemaMaster
        }
    }

    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestDomainControllers)) {
        # External command from PSSharedGoods
        $Data.ForestDomainControllers = Get-WinADForestControllers
    }
    # Forest Sites
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSites, [ActiveDirectory]::ForestSites1, [ActiveDirectory]::ForestSites2)) {
        Write-Verbose 'Getting forest information - Forest Sites'
        $Data.ForestSites = Get-WinADForestSites
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSites1)) {
        Write-Verbose 'Getting forest information - Forest Sites1'
        $Data.ForestSites1 = Get-WinADForestSites1 -ForestSites $Data.ForestSites
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSites2)) {
        Write-Verbose 'Getting forest information - Forest Sites2'
        $Data.ForestSites2 = Get-WinADForestSites2 -ForestSites $Data.ForestSites
    }
    ## Forest Subnets
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSubnet , [ActiveDirectory]::ForestSubnets1, [ActiveDirectory]::ForestSubnets2)) {
        Write-Verbose 'Getting forest information - Forest Subnets'
        $Data.ForestSubnets = Get-WinADForestSubnets
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSubnets1)) {
        Write-Verbose 'Getting forest information - Forest Subnets1'
        $Data.ForestSubnets1 = Get-WinADForestSubnets1 -ForestSubnets $ForestSubnets
    }
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSubnets2)) {
        Write-Verbose 'Getting forest information - Forest Subnets2'
        $Data.ForestSubnets2 = Get-WinADForestSubnets2 -ForestSubnets $ForestSubnets
    }
    ## Forest Site Links
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestSiteLinks)) {
        Write-Verbose 'Getting forest information - Forest SiteLinks'
        $Data.ForestSiteLinks = Get-WinADForestSiteLinks
    }

    ## Forest Optional Features
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::ForestOptionalFeatures)) {
        Write-Verbose 'Getting forest information - Forest Optional Features'
        $Data.ForestOptionalFeatures = Get-WinADForestOptionalFeatures
    }
    $EndTime = Stop-TimeLog -Time $TimeToGenerate
    Write-Verbose "Getting forest information - Time to generate: $EndTime"
    ### Generate Data from Domains
    $Data.FoundDomains = [ordered]@{ }
    foreach ($Domain in $Data.Domains) {
        $Data.FoundDomains.$Domain = Get-WinADDomainInformation -Domain $Domain -TypesRequired $TypesRequired -PathToPasswords $PathToPasswords -PathToPasswordsHashes $PathToPasswordsHashes
    }
}
