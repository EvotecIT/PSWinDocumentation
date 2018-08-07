Function Get-GPOInfo {
    <#
        Source: https://gallery.technet.microsoft.com/Get-GPO-informations-b02e0fdf
    #>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory = $false)][ValidateScript( {Test-Connection $_ -Count 1 -Quiet})][String]$DomainName = $env:USERDNSDOMAIN
    )
    Begin {
        Write-Verbose -Message "Importing Group Policy module..."
        try {Import-Module -Name GroupPolicy -Verbose:$false -ErrorAction stop}
        catch {Write-Warning -Message "Failed to import GroupPolicy module"; continue}
    }
    Process {
        ForEach ($GPO in (Get-GPO -All -Domain $DomainName )) {
            Write-Verbose -Message "Processing $($GPO.DisplayName)..."
            [xml]$XmlGPReport = $GPO.generatereport('xml')
            #GPO version
            if ($XmlGPReport.GPO.Computer.VersionDirectory -eq 0 -and $XmlGPReport.GPO.Computer.VersionSysvol -eq 0) {$ComputerSettings = "NeverModified"}else {$ComputerSettings = "Modified"}
            if ($XmlGPReport.GPO.User.VersionDirectory -eq 0 -and $XmlGPReport.GPO.User.VersionSysvol -eq 0) {$UserSettings = "NeverModified"}else {$UserSettings = "Modified"}
            #GPO content
            if ($XmlGPReport.GPO.User.ExtensionData -eq $null) {$UserSettingsConfigured = $false}else {$UserSettingsConfigured = $true}
            if ($XmlGPReport.GPO.Computer.ExtensionData -eq $null) {$ComputerSettingsConfigured = $false}else {$ComputerSettingsConfigured = $true}
            #Output
            [ordered] @{
                'Name'                   = $XmlGPReport.GPO.Name
                'Links'                  = $XmlGPReport.GPO.LinksTo | Select-Object -ExpandProperty SOMPath
                'Has Computer Settings'  = $ComputerSettingsConfigured
                'Has User Settings'      = $UserSettingsConfigured
                'User Enabled'           = $XmlGPReport.GPO.User.Enabled
                'Computer Enabled'       = $XmlGPReport.GPO.Computer.Enabled
                'Computer Settings'      = $ComputerSettings
                'User Settings'          = $UserSettings
                'Gpo Status'             = $GPO.GpoStatus
                'Creation Time'          = $GPO.CreationTime
                'Modification Time'      = $GPO.ModificationTime
                'WMI Filter'             = $GPO.WmiFilter.name
                'WMI Filter Description' = $GPO.WmiFilter.Description
                'Path'                   = $GPO.Path
                'GUID'                   = $GPO.Id
                'SDDL'                   = $XmlGPReport.GPO.SecurityDescriptor.SDDL.'#text'
                'ACLs'                   = $XmlGPReport.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object -Process {
                    New-Object -TypeName PSObject -Property @{
                        'User'            = $_.trustee.name.'#Text'
                        'Permission Type' = $_.type.PermissionType
                        'Inherited'       = $_.Inherited
                        'Permissions'     = $_.Standard.GPOGroupedAccessEnum
                    }
                }
            }
        }
    }
    End {}
}

Get-GPOInfo | Select -First 1 | Format-TransposeTable | fl