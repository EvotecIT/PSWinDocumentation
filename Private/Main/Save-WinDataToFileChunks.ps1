function Save-WinDataToFileInChunks {
    param(
        [nullable[bool]] $Export,
        [string] $Type,
        [Object] $Data,
        [string] $FolderPath,
        [switch] $IsOffline,
        [string] $FileType = 'XML'
    )

    foreach ($Key in $Data.Keys) {
        $FilePath = [IO.Path]::Combine($FolderPath, "$Key.xml")
        Save-WinDataToFile -Export $Export -Type $Type -IsOffline:$IsOffline -Data $Data.$Key -FilePath $FilePath -FileType $FileType
    }
}