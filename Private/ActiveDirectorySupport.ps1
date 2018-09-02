function Get-ADObjectFromDistingusishedName {
    [CmdletBinding()]
    param (
        [string[]] $DistinguishedName,
        [Object[]] $ADCatalog,
        [string] $Type = '',
        [string] $Splitter # ', ' # Alternative for example [System.Environment]::NewLine
    )
    $FoundObjects = @()

    foreach ($Catalog in $ADCatalog) {
        foreach ($Object in $DistinguishedName) {
            $ADObject = $Catalog | Where { $_.DistinguishedName -eq $Object }
            if ($ADObject) {
                if ($Type -eq '') {
                    #Write-Verbose 'Get-ADObjectFromDistingusishedName - Whole object'
                    $FoundObjects += $ADObject
                } else {
                    #Write-Verbose 'Get-ADObjectFromDistingusishedName - Part of object'
                    $FoundObjects += $ADObject.$Type
                }
            }
        }
    }
    if ($Splitter) {
        return ($FoundObjects | Sort-Object) -join $Splitter
    } else {
        return $FoundObjects | Sort-Object
    }
}
function Convert-ToDateTime {
    [CmdletBinding()]
    param (
        [string] $Timestring,
        [string] $Ignore = '*1601*'
    )
    Try {
        $DateTime = ([datetime]::FromFileTime($Timestring))
    } catch {
        $DateTime = $null
    }
    #Write-Verbose "Convert-ToDateTime: $DateTime"
    if ($DateTime -eq $null -or $DateTime -like $Ignore) {
        return $null
    } else {
        return $DateTime
    }
}

function Convert-ToTimeSpan {
    [CmdletBinding()]
    param (
        $StartTime,
        $EndTime
    )
    if ($StartTime -and $EndTime) {
        try {
            $TimeSpan = (NEW-TIMESPAN -Start (GET-DATE) -End ($EndTime))
        } catch {
            $TimeSpan = $null
        }
    }
    if ($TimeSpan -ne $null) {
        return $TimeSpan
    } else {
        return $null
    }
}
function Convert-TimeToDays {
    [CmdletBinding()]
    param (
        $StartTime,
        $EndTime,
        [string] $Ignore = '*1601*'
    )
    if ($StartTime -and $EndTime) {
        try {
            if ($StartTime -notlike $Ignore -and $EndTime -notlike $Ignore) {
                $Days = (NEW-TIMESPAN -Start (GET-DATE) -End ($EndTime)).Days
            } else {
                $Days = $null
            }
        } catch {
            $Days = $null
        }
    }
    return $Days
}