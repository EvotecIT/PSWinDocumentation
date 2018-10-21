function Get-ComputerWindowsUpdates {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )

    $Data = Get-hotfix -ComputerName $vComputerName | select Description , HotFixId , InstalledBy, InstalledOn, Caption
    return $Data
}