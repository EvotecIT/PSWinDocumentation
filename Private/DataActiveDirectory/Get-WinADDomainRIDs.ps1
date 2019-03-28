function Get-WinADDomainRIDs {
    [CmdletBinding()]
    param(
        [Microsoft.ActiveDirectory.Management.ADDomain] $DomainInformation,
        [string] $Domain
    )
    # Critical for RID Pool Depletion: https://blogs.technet.microsoft.com/askds/2011/09/12/managing-rid-pool-depletion/
    #$DomainInformation.GetType() | fv -List -Property *
    $Time = Start-TimeLog
    Write-Verbose "Getting domain information - $Domain DomainRIDs"

    if ($null -eq $DomainInformation) {
        $DomainInformation = Get-ADDomain -Server $Domain
    }
    $rID = [ordered] @{ }
    $rID.'rIDs Master' = $DomainInformation.RIDMaster

    $Property = get-adobject "cn=rid manager$,cn=system,$($DomainInformation.DistinguishedName)" -Property RidAvailablePool -Server $rID.'rIDs Master'
    [int32]$totalSIDS = $($Property.RidAvailablePool) / ([math]::Pow(2, 32))
    [int64]$temp64val = $totalSIDS * ([math]::Pow(2, 32))
    [int32]$currentRIDPoolCount = $($Property.RidAvailablePool) - $temp64val
    [int64]$RidsRemaining = $totalSIDS - $currentRIDPoolCount

    $Rid.'rIDs Available Pool' = $Property.RidAvailablePool
    $rID.'rIDs Total SIDs' = $totalSIDS
    $rID.'rIDs Issued' = $CurrentRIDPoolCount
    $rID.'rIDs Remaining' = $RidsRemaining
    $rID.'rIDs Percentage' = if ($RidsRemaining -eq 0) { $RidsRemaining.ToString("P") } else { ($currentRIDPoolCount / $RidsRemaining * 100).ToString("P") }

    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose "Getting domain information - $Domain DomainRIDs Time: $EndTime"
    return $rID
}