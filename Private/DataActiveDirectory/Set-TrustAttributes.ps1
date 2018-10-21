Function Set-TrustAttributes {
    [cmdletbinding()]
    Param(
        [parameter(Mandatory = $false, ValueFromPipeline = $True)][int32]$Value
    )
    [String[]]$TrustAttributes = @()
    Foreach ($V in $Value) {
        if ([int32]$V -band 0x00000001) {$TrustAttributes += "Non Transitive"}
        if ([int32]$V -band 0x00000002) {$TrustAttributes += "UpLevel"}
        if ([int32]$V -band 0x00000004) {$TrustAttributes += "Quarantaine (SID Filtering enabled)"} #SID Filtering
        if ([int32]$V -band 0x00000008) {$TrustAttributes += "Forest Transitive"}
        if ([int32]$V -band 0x00000010) {$TrustAttributes += "Cross Organization (Selective Authentication enabled)"} #Selective Auth
        if ([int32]$V -band 0x00000020) {$TrustAttributes += "Within Forest"}
        if ([int32]$V -band 0x00000040) {$TrustAttributes += "Treat as External"}
        if ([int32]$V -band 0x00000080) {$TrustAttributes += "Uses RC4 Encryption"}
    }
    return $TrustAttributes
}