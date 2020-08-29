function New-HTMLGroupDiagramHierachical {
    [cmdletBinding()]
    param(
        [Array] $ADGroup,
        [ValidateSet('Default', 'Hierarchical', 'Both')][string] $RemoveAppliesTo = 'Both',
        [switch] $RemoveComputers,
        [switch] $RemoveUsers,
        [switch] $RemoveOther
    )
    New-HTMLDiagram -Height 'calc(100vh - 400px)' {
        New-DiagramOptionsLayout -HierarchicalEnabled $true #-HierarchicalDirection FromLeftToRight #-HierarchicalSortMethod directed
        New-DiagramOptionsPhysics -Enabled $true -HierarchicalRepulsionAvoidOverlap 1 -HierarchicalRepulsionNodeDistance 150
        #New-DiagramOptionsPhysics -RepulsionNodeDistance 150 -Solver repulsion
        if ($ADGroup) {
            # Add it's members to diagram
            foreach ($ADObject in $ADGroup) {
                # Lets build our diagram
                [int] $Level = $($ADObject.Nesting) + 1
                $ID = "$($ADObject.DomainName)$($ADObject.Name)$Level"
                [int] $LevelParent = $($ADObject.Nesting)
                $IDParent = "$($ADObject.ParentGroupDomain)$($ADObject.ParentGroup)$LevelParent"

                [int] $Level = $($ADObject.Nesting) + 1
                if ($ADObject.Type -eq 'User') {
                    if (-not $RemoveUsers -or $RemoveAppliesTo -notin 'Both', 'Hierarchical') {
                        $Label = $ADObject.Name + [System.Environment]::NewLine + $ADObject.DomainName
                        New-DiagramNode -Id $ID -Label $Label -To $IDParent -Image 'https://image.flaticon.com/icons/svg/3135/3135715.svg' -Level $Level -ArrowsToEnabled
                    }
                } elseif ($ADObject.Type -eq 'Group') {
                    if ($ADObject.Nesting -eq -1) {
                        $BorderColor = 'Red'
                        $Image = 'https://image.flaticon.com/icons/svg/921/921347.svg'
                    } else {
                        $BorderColor = 'Blue'
                        $Image = 'https://image.flaticon.com/icons/svg/166/166258.svg'
                    }
                    $SummaryMembers = -join ('Total: ', $ADObject.TotalMembers, ' Direct: ', $ADObject.DirectMembers, ' Groups: ', $ADObject.DirectGroups, ' Indirect: ', $ADObject.IndirectMembers)
                    $Label = $ADObject.Name + [System.Environment]::NewLine + $ADObject.DomainName + [System.Environment]::NewLine + $SummaryMembers
                    New-DiagramNode -Id $ID -Label $Label -To $IDParent -Image $Image -Level $Level -ArrowsToEnabled -LinkColor Orange -ColorBorder $BorderColor
                } elseif ($ADObject.Type -eq 'Computer') {
                    if (-not $RemoveComputers -or $RemoveAppliesTo -notin 'Both', 'Hierarchical') {
                        $Label = $ADObject.Name + [System.Environment]::NewLine + $ADObject.DomainName
                        New-DiagramNode -Id $ID -Label $Label -To $IDParent -Image 'https://image.flaticon.com/icons/svg/3003/3003040.svg' -Level $Level -ArrowsToEnabled
                    }
                } else {
                    if (-not $RemoveOther -or $RemoveAppliesTo -notin 'Both', 'Hierarchical') {
                        $Label = $ADObject.Name + [System.Environment]::NewLine + $ADObject.DomainName
                        New-DiagramNode -Id $ID -Label $Label -To $IDParent -Image 'https://image.flaticon.com/icons/svg/3003/3003040.svg' -Level $Level
                    }
                }
            }
        }
    }
}