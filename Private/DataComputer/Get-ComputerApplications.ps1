function Get-ComputerApplications {
    param(
        $ComputerName = $Env:COMPUTERNAME
    )
    $ScriptBlock = {
        $objapp1 = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*
        $objapp2 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*

        $app1 = $objapp1 | Select-Object Displayname, Displayversion , Publisher, Installdate, @{Expression = { 'x64' }; Label = "WindowsType"}
        $app2 = $objapp2 | Select-Object Displayname, Displayversion , Publisher, Installdate, @{Expression = { 'x86' }; Label = "WindowsType"} | where { -NOT (([string]$_.displayname).contains("Security Update for Microsoft") -or ([string]$_.displayname).contains("Update for Microsoft"))}
        $app = $app1 + $app2 #| Sort-Object -Unique
        return $app | Where { $_.Displayname -ne $null } | Sort-Object DisplayName
    }
    if ($ComputerName -eq $Env:COMPUTERNAME) {
        $Data = Invoke-Command -ScriptBlock $ScriptBlock
    } else {
        $Data = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock
    }
    return $Data

}