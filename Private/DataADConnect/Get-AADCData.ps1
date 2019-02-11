function Get-AADCData {
    param(
        [string] $ComputerName
    )
    $AADConnect = Invoke-Command -ComputerName $ComputerName -ScriptBlock {        
        $Data = @{}
        $Data.AutoUpgrade = Get-ADSyncAutoUpgrade
        $Data.GlobalSettings = Invoke-Command -ScriptBlock {
            $Settings = Get-ADSyncGlobalSettings      
            [PSCustomObject] @{
                FormatVersion       = $Settings.FormatVersion
                Version             = $Settings.Version
                PasswordSyncEnabled = $Settings.PasswordSyncEnabled
                NumSavedPwdEvents   = $Settings.NumSavedPwdEvents
                SchemaGuid          = $Settings.Schema.Identifier.Guid
            }         
        }
        $Data.GlobalSettingsParameters = Invoke-Command -ScriptBlock {
            $Global = Get-ADSyncGlobalSettings      
            $Global.Parameters          
        }
        $Data.GlobalSettingsObjects = Invoke-Command -ScriptBlock {
            $ObjectTypes = (Get-ADSyncGlobalSettings).Schema.ObjectTypes

            $Object = @{}
            foreach ($Type in $ObjectTypes) {
                $Object."$($Type.Identifier)" = $Type
            }

            #[PSCustomObject] @{

            #}       
        }
        $Data.GlobalSettingsTypes = Invoke-Command -ScriptBlock {
            $Settings = (Get-ADSyncGlobalSettings).Schema.AttributeTypes
            $Settings
            #[PSCustomObject] @{
                
            #}       
        }
        $Data.Schedule = Get-ADSyncScheduler

        $Data.Rules = Invoke-Command -ScriptBlock {
            $Rules = Get-ADSyncRule
            foreach ($Rule in $Rules) {
                [PSCustomObject] @{
                    Identifier                             = $Rule.Identifier
                    Name                                   = $Rule.Name
                    Version                                = $Rule.Version
                    Description                            = $Rule.Description
                    # it's already part of main object 
                    #ImmutableTag                        = $Rule.ImmutableTag                  
                    Connector                              = $Rule.Connector
                    Direction                              = $Rule.Direction
                    Disabled                               = $Rule.Disabled
                    SourceObjectType                       = $Rule.SourceObjectType
                    TargetObjectType                       = $Rule.TargetObjectType
                    Precedence                             = $Rule.Precedence
                    PrecedenceAfter                        = $Rule.PrecedenceAfter
                    PrecedenceBefore                       = $Rule.PrecedenceBefore
                    LinkType                               = $Rule.LinkType
                    EnablePasswordSync                     = $Rule.EnablePasswordSync
                 
                    
                    JoinFilterConditionCSAttribute         = $Rule.JoinFilter.JoinConditionList.CSAttribute -join ','
                    JoinFilterConditionMVAttribute         = $Rule.JoinFilter.JoinConditionList.MVAttribute -join ','
                    JoinFilterConditionCaseSensitive       = $Rule.JoinFilter.JoinConditionList.CaseSensitive -join ','      
                    JoinFilterHash                         = $Rule.JoinFilter.JoinHash -join ','
                    
                    ScopeFilterConditionAttribute          = $Rule.ScopeFilter.ScopeConditonList.Attribute
                    ScopeFilterConditionComparisonValue    = $Rule.ScopeFilter.ScopeConditonList.ComparisonValue
                    ScopeFilterConditionComparisonOperator = $Rule.ScopeFilter.ScopeConditonList.ComparisonOperator           
                    
                    AttributeFlowMappingsSource            = $Rule.AttributeFlowMappings.Source -join ','
                    AttributeFlowMappingsDestination       = $Rule.AttributeFlowMappings.Destination
                    AttributeFlowMappingsFlowType          = $Rule.AttributeFlowMappings.FlowType
                    AttributeFlowMappingsExecuteOnce       = $Rule.AttributeFlowMappings.ExecuteOnce
                    AttributeFlowMappingsExpression        = $Rule.AttributeFlowMappings.Expression
                    AttributeFlowMappingsValueMergeType    = $Rule.AttributeFlowMappings.ValueMergeType
                    AttributeFlowMappingsMappingSource     = $Rule.AttributeFlowMappings.MappingSourceAsString

                    SourceNamespaceId                      = $Rule.SourceNamespaceId
                    TargetNamespaceId                      = $Rule.TargetNamespaceId
                    VersionAgnosticTag                     = $Rule.VersionAgnosticTag
                    TagVersion                             = $Rule.TagVersion
                    IsStandardRule                         = $Rule.IsStandardRule
                    IsLegacyCustomRule                     = $Rule.IsLegacyCustomRole
                    JoinHash                               = $Rule.JoinHash                   

                }
            }
        }
        $Data.Connector = Invoke-Command -ScriptBlock {
            $Connector = Get-ADSyncConnector

            foreach ($Connect in $Connector) {
                [PsCustomObject] @{

                }
            }

        }
        return $Data
    }
    $AADConnect
}

#$AADC = Get-AADCData -ComputerName 'ADConnect'
#$AADC.Rules.AttributeFlowMappingsExpression | Format-Table *
