function Test-File {
    [CmdletBinding()]
    param (
        [string] $File,
        [string] $FileName,
        [switch] $Require,
        [switch] $Skip
    )
    [int] $ErrorCount = 0
    if ($Skip) {
        return $ErrorCount
    }
    if ($File -ne '') {
        if ($Require) {
            if (Test-Path $File) {
                return $ErrorCount
            } else {
                Write-Color  @Script:WriteParameters '[e] ', $FileName, " doesn't exists (", $File, "). It's required if you want to use this feature." -Color Red, Yellow, Yellow, White
                $ErrorCount++
            }
        }
    } else {
        $ErrorCount++
        Write-Color @Script:WriteParameters '[e] ', $FileName, " was empty. It's required if you want to use this feature." -Color Red, Yellow, White
    }
    return $ErrorCount
}
