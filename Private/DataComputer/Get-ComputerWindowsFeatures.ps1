function Get-ComputerWindowsFeatures {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data = Get-WmiObject Win32_OptionalFeature -ComputerName $vComputerName | select Caption , Installstate
    return $Data
}