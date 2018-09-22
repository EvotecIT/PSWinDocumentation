function Search-Command {
    [cmdletbinding()]
    param (
        $CommandName
    )
    return [bool](Get-Command -Name $CommandName -ErrorAction SilentlyContinue)
}