function Test-File {
    [CmdletBinding()]
    param (
        [string] $File,
        [string] $FileName,
        [switch] $Require,
        [switch] $Skip
    )
    [int] $ErrorCount = 0
    if ($Skip) {
        return $ErrorCount
    }
    if ($File -ne '') {
        if ($Require) {
            if (Test-Path $File) {
                return $ErrorCount
            } else {
                Write-Color  @Script:WriteParameters '[e] ', $FileName, " doesn't exists (", $File, "). It's required if you want to use this feature." -Color Red, Yellow, Yellow, White
                $ErrorCount++
            }
        }
    } else {
        $ErrorCount++
        Write-Color @Script:WriteParameters '[e] ', $FileName, " was empty. It's required if you want to use this feature." -Color Red, Yellow, White
    }
    return $ErrorCount
}

function Test-Configuration {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    [int] $ErrorCount = 0
    $Script:WriteParameters = $Document.Configuration.DisplayConsole


    $Keys = Get-ObjectKeys -Object $Document -Ignore 'Configuration'
    foreach ($Key in $Keys) {
        $ErrorCount += Test-File -File $Document.$Key.FilePathWord -FileName 'FilePathWord' -Skip:(-not $Document.$Key.ExportWord)
        $ErrorCount += Test-File -File $Document.$Key.FilePathExcel -FileName 'FilePathExcel' -Skip:(-not $Document.$Key.ExportExcel)
    }
    if ($ErrorCount -ne 0) {
        Exit
    }
}
Function Get-Types {
    [CmdletBinding()]
    param (
        $Types = [ActiveDirectory]
    )
    $TypesRequired = @()
    foreach ($Type in $Types) {
        $TypesRequired += $Type.GetEnumValues()
    }
    return $TypesRequired
}

function Get-TypesRequired {
    [CmdletBinding()]
    param (
        [System.Object] $Document
    )
    $ADSectionsForest = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionForest
    $ADSectionsDomain = Get-ObjectKeys -Object $Document.DocumentAD.Sections.SectionDomain

    $TypesRequired = New-ArrayList
    foreach ($Section in $ADSectionsDomain) {
        if ($Document.DocumentAD.Sections.SectionDomain.$Section.Use -eq $True) {
            Add-ToArrayAdvanced -List $TypesRequired -Element $Document.DocumentAD.Sections.SectionDomain.$Section.TableData -SkipNull -RequireUnique -FullComparison
            Add-ToArrayAdvanced -List $TypesRequired -Element $Document.DocumentAD.Sections.SectionDomain.$Section.ListData -SkipNull -RequireUnique -FullComparison
            Add-ToArrayAdvanced -List $TypesRequired -Element $Document.DocumentAD.Sections.SectionDomain.$Section.ChartData -SkipNull -RequireUnique -FullComparison
        }
    }

    foreach ($Section in $ADSectionsForest) {
        if ($Document.DocumentAD.Sections.SectionForest.$Section.Use -eq $True) {
            Add-ToArrayAdvanced -List $TypesRequired -Element $Document.DocumentAD.Sections.SectionForest.$Section.TableData -SkipNull -RequireUnique -FullComparison
            Add-ToArrayAdvanced -List $TypesRequired -Element $Document.DocumentAD.Sections.SectionForest.$Section.ListData -SkipNull -RequireUnique -FullComparison
            Add-ToArrayAdvanced -List $TypesRequired -Element $Document.DocumentAD.Sections.SectionForest.$Section.ChartData -SkipNull -RequireUnique -FullComparison
        }
    }
    #Show-Array $TypesRequired -WithType
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
    $Test = Test-ComputerSecureChannel
    if ($Test) {
        # future use
    } else {
        Write-Warning 'No connectivity to forest/domain.'
        Exit
    }
}