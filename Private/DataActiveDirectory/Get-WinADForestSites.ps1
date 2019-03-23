function Get-WinADForestSites {
    [CmdletBinding()]
    param(

    )
    $ExludedProperties = @(
        'PropertyNames', 'AddedProperties', 'RemovedProperties', 'ModifiedProperties', 'PropertyCount'
    )

    $Properties = @('Name',
        'DisplayName', 'Description', 'CanonicalName', 'DistinguishedName', 'Location', 'ManagedBy', 'Created', 'Modified', 'Deleted',
        'ProtectedFromAccidentalDeletion', 'RedundantServerTopologyEnabled',
        'AutomaticInterSiteTopologyGenerationEnabled',
        'AutomaticTopologyGenerationEnabled',
        'Subnets',
        #'nTSecurityDescriptor'
        #LastKnownParent
        #instanceType
        #InterSiteTopologyGenerator
        #dSCorePropagationData
        #ReplicationSchedule.RawSchedule -join ','
        #msExchServerSiteBL -join ','
        #siteObjectBL -join ','
        #systemFlags
        #ObjectGUID
        #ObjectCategory
        #ObjectClass
        #ScheduleHashingEnabled
        'sDRightsEffective',
        'TopologyCleanupEnabled',
        'TopologyDetectStaleEnabled',
        'TopologyMinimumHopsEnabled',
        'UniversalGroupCachingEnabled',
        'UniversalGroupCachingRefreshSite',
        'WindowsServer2000BridgeheadSelectionMethodEnabled',
        'WindowsServer2000KCCISTGSelectionBehaviorEnabled',
        'WindowsServer2003KCCBehaviorEnabled',
        'WindowsServer2003KCCIgnoreScheduleEnabled',
        'WindowsServer2003KCCSiteLinkBridgingEnabled'
    )
    return Get-ADReplicationSite -Filter * -Properties $Properties | Select-Object -Property $Properties -ExcludeProperty $ExludedProperties
}









