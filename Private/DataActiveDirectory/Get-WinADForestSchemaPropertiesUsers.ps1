function Get-WinADForestSchemaPropertiesUsers {
    [CmdletBinding()]
    param(

    )
    Write-Verbose "Getting forest information - ForestSchemaPropertiesUsers"
    $Time = Start-TimeLog
    $Schema = [directoryservices.activedirectory.activedirectoryschema]::GetCurrentSchema()
    @(
        $Schema.FindClass("user").mandatoryproperties | Select-Object name, commonname, description, syntax #| export-csv user-mandatory-attributes.csv -Delimiter ';'
        $Schema.FindClass("user").optionalproperties | Select-Object name, commonname, description, syntax #| export-csv user-optional-attributes.csv -Delimiter ';'
    )
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose "Getting domain information - ForestSchemaPropertiesUsers Time: $EndTime"
}