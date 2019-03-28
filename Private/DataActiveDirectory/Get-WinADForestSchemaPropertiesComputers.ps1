function Get-WinADForestSchemaPropertiesComputers {
    [CmdletBinding()]
    param(

    )
    Write-Verbose "Getting forest information - ForestSchemaPropertiesComputers"
    $Time = Start-TimeLog
    $Schema = [directoryservices.activedirectory.activedirectoryschema]::GetCurrentSchema()
    @(
        $Schema.FindClass("computer").mandatoryproperties | Select-Object name, commonname, description, syntax
        $Schema.FindClass("computer").optionalproperties | Select-Object name, commonname, description, syntax #| Where-Object { $_.Name -eq 'ms-Mcs-AdmPwd' } # ft -AutoSize
    )
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose "Getting domain information - ForestSchemaPropertiesComputers Time: $EndTime"
}


<#
Get-WinADForestSchemaProperties -Verbose


function Get-WinADForestSchemaProperties {
    [CmdletBinding()]
    param(

    )
    Write-Verbose "Getting forest information - ForestSchemaProperties"
    $Time = Start-TimeLog

    $Output = @{}
    $Schema = [directoryservices.activedirectory.activedirectoryschema]::GetCurrentSchema()
    $Output.SchemaPropertiesUser = @(
        $Mandatory = $Schema.FindClass("user").mandatoryproperties #| Select-Object name, commonname, description, syntax #| export-csv user-mandatory-attributes.csv -Delimiter ';'
        foreach ($Object in $Mandatory) {
            [PSCustomobject] @{
                Name        = $Object.Name
                CommonName  = $Object.CommonName
                Description = $Object.Description
                Syntax      = $Object.Syntax
            }
        }
        $Optional = $Schema.FindClass("user").optionalproperties # | Select-Object name, commonname, description, syntax #| export-csv user-optional-attributes.csv -Delimiter ';'
        foreach ($Object in $Optional) {
            [PSCustomobject] @{
                Name        = $Object.Name
                CommonName  = $Object.CommonName
                Description = $Object.Description
                Syntax      = $Object.Syntax
            }
        }
    )
    $Output.SchemaPropertiesComputer = @(
        $Mandatory = $Schema.FindClass("computer").mandatoryproperties #| Select-Object name, commonname, description, syntax
        foreach ($Object in $Mandatory) {
            [PSCustomobject] @{
                Name        = $Object.Name
                CommonName  = $Object.CommonName
                Description = $Object.Description
                Syntax      = $Object.Syntax
            }
        }
        $Optional = $Schema.FindClass("computer").optionalproperties #| Select-Object name, commonname, description, syntax #| Where-Object { $_.Name -eq 'ms-Mcs-AdmPwd' } # ft -AutoSize
        foreach ($Object in $Optional) {
            [PSCustomobject] @{
                Name        = $Object.Name
                CommonName  = $Object.CommonName
                Description = $Object.Description
                Syntax      = $Object.Syntax
            }
        }
    )
    $EndTime = Stop-TimeLog -Time $Time -Option OneLiner
    Write-Verbose "Getting domain information - ForestSchemaProperties Time: $EndTime"
    return $Output
}

Get-WinADForestSchemaProperties -Verbose

#>