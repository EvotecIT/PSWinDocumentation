function Get-ComputerCulture {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $ScriptBlock = {
        get-culture | select KeyboardLayoutId, DisplayName, @{Expression = {$_.ThreeLetterWindowsLanguageName}; Label = "Windows Language"}
    }
    if ($ComputerName -eq $Env:COMPUTERNAME) {
        $Data8 = Invoke-Command -ScriptBlock $ScriptBlock
    } else {
        $Data8 = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock
    }
    return $Data8
}