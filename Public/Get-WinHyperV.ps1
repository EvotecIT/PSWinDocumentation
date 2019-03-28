function Get-WinHyperV {
    param(

    )
    # get-command -module hyper-v
    $Data = [ordered] @{}
    $Data.VM = Get-VM
    $Data.VMCheckpoints = Get-VMCheckpoint ($Data.VM)
    $Data.VMFirmware =  Get-VMFirmware AD1
    $Data.VMBios = Get-VMBios AD1 # may not work, above one works

    #Get-VMBios : The Generation 2 virtual machine or snapshot "AD1" does not support the VMBios cmdlets. Use Get-VMFirmware and Set-VMFirmware instead.
}