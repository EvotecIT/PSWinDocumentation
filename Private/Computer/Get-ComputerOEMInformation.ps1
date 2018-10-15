function Get-ComputerOemInformation {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $ScriptBlock = { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation | select Model, Manufacturer, Logo, SupportPhone, SupportURL, SupportHours }
    if ($ComputerName -eq $Env:COMPUTERNAME) {
        $Data = Invoke-Command -ScriptBlock $ScriptBlock
    } else {
        $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock
    }
    return $Data
}