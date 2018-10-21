function Test-ForestConnectivity {

    Try {
        $Test = Get-ADForest
        return $true
    } catch {
        #Write-Warning 'No connectivity to forest/domain.'
        return $False
    }
}