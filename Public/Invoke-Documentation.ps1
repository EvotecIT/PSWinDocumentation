function Invoke-Documentation {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('ActiveDirectory')][string] $Service,
        [Parameter()][PSWinDocumentation.ActiveDirectory[]] $ActiveDirectoryServices,
        [string] $PasswordFile,
        [string] $PasswordFileHashes,
        [Parameter(Mandatory)][ValidateSet('HTML', 'Word', 'Excel')][string[]] $Output,
        [string] $FilePath
    )
    Begin {
        # [System.IO.Path] | Get-Member -Static
        if ($FilePath) {
            $DirectoryName = [System.io.path]::GetDirectoryName($FilePath)
            $FilePathWithOutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
            Write-Color -Text "FilePath was given. Saving file to:"
            if ($Output -contains 'Word') {
                $FilePathWord = [IO.Path]::Combine($DirectoryName, "$($FilePathWithOutExtension).docx")
                Write-Color -Text "[x] Word: $FilePathWord"
            }
            if ($Output -contains 'HTML') {
                $FilePathHTML = [IO.Path]::Combine($DirectoryName, "$($FilePathWithOutExtension).html")
                Write-Color -Text "[x] HTML: $FilePathHTML"
            }
            if ($Output -contains 'Excel') {
                $FilePathExcel = [IO.Path]::Combine($DirectoryName, "$($FilePathWithOutExtension).xlsx")
                Write-Color -Text "[x] Excel: $FilePathExcel"
            }

        } else {
            $DesktopPath = [Environment]::GetFolderPath("Desktop")
            Write-Color -Text "FilePath was not given. Using defaults. Saving file to:" -Color Yellow
            if ($Output -contains 'Word') {
                $FilePathWord = [IO.Path]::Combine($DesktopPath, "PSWinDocumentation-$((Get-Date).ToString('yyyy-MM-dd_HH_mm_ss')).docx")
                Write-Color -Text "[x] Word: $FilePathWord" -Color Yellow
            }
            if ($Output -contains 'HTML') {

            }
            if ($Output -contains 'Excel') {

            }
        }
    }
    Process {
        if ($Service -eq 'ActiveDirectory') {
            if ($null -eq $DataSetForest) {
                $DataSetForest = Get-WinADForestInformation -PasswordQuality -DontRemoveEmpty -PathToPasswords $PasswordFile -PathToPasswordsHashes $PasswordFileHashes
            }
            if ($Output -contains 'Word') {
                Invoke-ADWord -FilePath $FilePathWord -DataSetForest $DataSetForest
            }
            if ($Output -contains 'HTML') {
                Invoke-ADHTML -FilePath $FilePathHTML -DataSetForest $DataSetForest
            }
            if ($Output -contains 'Excel') {
                Invoke-ADExcel -FilePath $FilePathExcel -DataSetForest $DataSetForest
            }
        } elseif ($Service -eq 'O365') {

        }
    }
}