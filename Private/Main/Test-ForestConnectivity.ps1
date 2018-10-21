function Test-ForestConnectivity {
    Try {
        $Test = Get-ADForest
    } catch {
        Write-Warning 'No connectivity to forest/domain.'
        Exit
    }
}