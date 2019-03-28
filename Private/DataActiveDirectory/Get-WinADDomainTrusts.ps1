function Get-WinADDomainTrusts {
    [CmdletBinding()]
    param(
        [string] $Domain,
        [string] $DomainPDC,
        [Array] $Trusts,
        [Array] $TypesRequired
    )
    if (Find-TypesNeeded -TypesRequired $TypesRequired -TypesNeeded @([ActiveDirectory]::DomainTrusts)) {
        Write-Verbose "Getting domain information - $Domain DomainTrusts"
        $Time = Start-TimeLog

        if ($null -eq $Trusts) {
            $Trusts = Get-ADTrust -Server $Domain -Filter * -Properties *
        }
        if ($DomainPDC -eq '') {
            $DomainPDC = (Get-ADDomain -Server $Domain).PDCEmulator
        }

        $PropertiesTrustWMI = @(
            'FlatName',
            'SID',
            'TrustAttributes',
            'TrustDirection',
            'TrustedDCName',
            'TrustedDomain',
            'TrustIsOk',
            'TrustStatus',
            'TrustStatusString', # TrustIsOk/TrustStatus are covered by this
            'TrustType'
        )

        <# TrustWMI
        FlatName          : EVOTECPL
        SID               : S-1-5-21-3661168273-3802070955-2987026695
        TrustAttributes   : 32
        TrustDirection    : 3
        TrustedDCName     : \\ADPreview2019.ad.evotec.pl
        TrustedDomain     : ad.evotec.pl
        TrustIsOk         : True
        TrustStatus       : 0
        TrustStatusString : OK
        TrustType         : 2
        PSComputerName    : ad1.ad.evotec.xyz
    #>

        $TrustStatatuses = Get-CimInstance -ClassName Microsoft_DomainTrustStatus -Namespace root\MicrosoftActiveDirectory -ComputerName $DomainPDC -ErrorAction SilentlyContinue -Verbose:$false -Property $PropertiesTrustWMI

        $ReturnData = foreach ($Trust in $Trusts) {
            $TrustWMI = $TrustStatatuses | & { process { if ($_.TrustedDomain -eq $Trust.Target ) { $_ } } }
            [PSCustomObject][ordered] @{
                'Trust Source'               = $Domain
                'Trust Target'               = $Trust.Target
                'Trust Direction'            = $Trust.Direction
                'Trust Attributes'           = if ($Trust.TrustAttributes -is [int]) { Set-TrustAttributes -Value $Trust.TrustAttributes } else { 'Error - needs fixing' }
                'Trust Status'               = if ($null -ne $TrustWMI) { $TrustWMI.TrustStatusString } else { 'N/A' }
                'Forest Transitive'          = $Trust.ForestTransitive
                'Selective Authentication'   = $Trust.SelectiveAuthentication
                'SID Filtering Forest Aware' = $Trust.SIDFilteringForestAware
                'SID Filtering Quarantined'  = $Trust.SIDFilteringQuarantined
                'Disallow Transivity'        = $Trust.DisallowTransivity
                'Intra Forest'               = $Trust.IntraForest
                'Tree Parent?'               = $Trust.IsTreeParent
                'Tree Root?'                 = $Trust.IsTreeRoot
                'TGTDelegation'              = $Trust.TGTDelegation
                'TrustedPolicy'              = $Trust.TrustedPolicy
                'TrustingPolicy'             = $Trust.TrustingPolicy
                'TrustType'                  = $Trust.TrustType
                'UplevelOnly'                = $Trust.UplevelOnly
                'UsesAESKeys'                = $Trust.UsesAESKeys
                'UsesRC4Encryption'          = $Trust.UsesRC4Encryption
                'Trust Source DC'            = if ($null -ne $TrustWMI) { $TrustWMI.PSComputerName } else { 'N/A' }
                'Trust Target DC'            = if ($null -ne $TrustWMI) { $TrustWMI.TrustedDCName.Replace('\\', '') } else { 'N/A' }
                'Trust Source DN'            = $Trust.Source
                'ObjectGUID'                 = $Trust.ObjectGUID
                'Created'                    = $Trust.Created
                'Modified'                   = $Trust.Modified
                'Deleted'                    = $Trust.Deleted
                'SID'                        = $Trust.securityIdentifier
                'TrustOK'                    = if ($null -ne $TrustWMI) { $TrustWMI.TrustIsOK } else { $false }
                'TrustStatus'                = if ($null -ne $TrustWMI) { $TrustWMI.TrustStatus } else { -1 }
            }
        }

        $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
        Write-Verbose "Getting domain information - $Domain DomainTrusts Time: $EndTime"

        return $ReturnData
    }
}