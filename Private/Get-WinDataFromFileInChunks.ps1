function Get-WinDataFromFileInChunks {
    [CmdletBinding()]
    param (
        [string] $FolderPath,
        [string] $FileType = 'XML',
        [Object] $Type
    )
    $DataInformation = @{}
    if (Test-Path $FolderPath) {
        $Files = @( Get-ChildItem -Path "$FolderPath\*.$FileType" -ErrorAction SilentlyContinue -Recurse )
        foreach ($File in $Files) {
            $FilePath = $File.FullName
            $FieldName = $File.BaseName
            Write-Verbose -Message "Importing $FilePath as $FieldName"
            try {
                $DataInformation.$FieldName = Import-CliXML -Path $FilePath -ErrorAction Stop
            } catch {
                $ErrorMessage = $_.Exception.Message -replace "`n", " " -replace "`r", " "
                Write-Warning "Couldn't load $FileType file from $FilePath for $Type data to match into $FieldName. Error occured: $ErrorMessage"
            }
        }
    } else {
        Write-Warning -Message "Couldn't load files ($FileType) from folder $FolderPath as it doesn't exists."
    }
    return $DataInformation
}
<# Simple Use Case

$Data = Get-WinDataFromFileInChunks -FolderPath "$Env:USERPROFILE\Desktop\PSWinDocumentation"
$Data | Format-Table -AutoSize
$Data.FoundDomains.'ad.evotec.xyz' | Ft -a

#>