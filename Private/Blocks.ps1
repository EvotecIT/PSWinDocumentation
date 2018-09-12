Function Get-Types {
    [CmdletBinding()]
    param (
        [Object] $Types
    )
    $TypesRequired = @()
    foreach ($Type in $Types) {
        #Write-Verbose "Type: $Type"
        $TypesRequired += $Type.GetEnumValues()
    }
    return $TypesRequired
}

function Get-TypesRequired {
    [CmdletBinding()]
    param (
        [hashtable[]] $Sections
    )
    $TypesRequired = New-ArrayList
    $Types = 'TableData', 'ListData', 'ChartData', 'SqlData', 'ExcelData'
    foreach ($Section in $Sections) {
        $Keys = Get-ObjectKeys -Object $Section
        foreach ($Key in $Keys) {
            if ($Section.$Key.Use -eq $True) {
                foreach ($Type in $Types) {
                    #Write-Verbose "Get-TypesRequired - Section: $Key Type: $Type Value: $($Section.$Key.$Type)"
                    Add-ToArrayAdvanced -List $TypesRequired -Element $Section.$Key.$Type -SkipNull -RequireUnique -FullComparison
                }
            }
        }
    }
    Write-Verbose "Get-TypesRequired - FinalList: $($TypesRequired -join ' ,')"
    return $TypesRequired
}
function Get-DocumentPath {
    [CmdletBinding()]
    param (
        [System.Object] $Document,
        [string] $FinalDocumentLocation
    )
    if ($Document.Configuration.Prettify.UseBuiltinTemplate) {
        #Write-Verbose 'Get-DocumentPath - Option 1'
        $WordDocument = Get-WordDocument -FilePath "$((get-item $PSScriptRoot).Parent.FullName)\Templates\WordTemplate.docx"
    } else {
        if ($Document.Configuration.Prettify.CustomTemplatePath) {
            if (Test-File -File $Document.Configuration.Prettify.CustomTemplatePath -FileName 'CustomTemplatePath' -eq 0) {
                # Write-Verbose 'Get-DocumentPath - Option 2'
                $WordDocument = Get-WordDocument -FilePath $Document.Configuration.Prettify.CustomTemplatePath
            } else {
                #Write-Verbose 'Get-DocumentPath - Option 3'
                $WordDocument = New-WordDocument -FilePath $FinalDocumentLocation
            }
        } else {
            #Write-Verbose 'Get-DocumentPath - Option 4'
            $WordDocument = New-WordDocument -FilePath $FinalDocumentLocation
        }
    }
    if ($WordDocument -eq $null) { Write-Verbose ' Null'}
    return $WordDocument
}
function Search-Command($CommandName) {
    return [bool](Get-Command -Name $CommandName -ErrorAction SilentlyContinue)
}

function Test-ModuleAvailability {
    if (Search-Command -CommandName 'Get-AdForest') {
        # future use
    } else {
        Write-Warning 'Modules required to run not found.'
        Exit
    }
}
function Test-ForestConnectivity {
    Try {
        $Test = Get-ADForest
    } catch {
        Write-Warning 'No connectivity to forest/domain.'
        Exit
    }
}