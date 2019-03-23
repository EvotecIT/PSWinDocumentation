Function Set-TrustAttributes {
    [cmdletbinding()]
    Param(
        [parameter(Mandatory = $false, ValueFromPipeline = $True)][int32]$Value
    )
    [String[]]$TrustAttributes = @(
        Foreach ($V in $Value) {
            if ([int32]$V -band 0x00000001) { "Non Transitive" }
            if ([int32]$V -band 0x00000002) { "UpLevel" }
            if ([int32]$V -band 0x00000004) { "Quarantaine (SID Filtering enabled)" } #SID Filtering
            if ([int32]$V -band 0x00000008) { "Forest Transitive" }
            if ([int32]$V -band 0x00000010) { "Cross Organization (Selective Authentication enabled)" } #Selective Auth
            if ([int32]$V -band 0x00000020) { "Within Forest" }
            if ([int32]$V -band 0x00000040) { "Treat as External" }
            if ([int32]$V -band 0x00000080) { "Uses RC4 Encryption" }
        }
    )
    return $TrustAttributes
}