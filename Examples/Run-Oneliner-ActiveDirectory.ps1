Import-Module $PSScriptRoot\..\PSWinDocumentation.psd1 -Force
#Import-Module PSWinDocumentation.AD

Invoke-Documentation -Service ActiveDirectory -Output Word, HTML -FilePath $Env:USERPROFILE\Desktop\MyReport

<#
[System.io.path]::GetDirectoryName("$Env:USERPROFILE\Desktop\Test.docx")
[System.io.path]::GetFileNameWithoutExtension("$Env:USERPROFILE\Desktop\Test.docx")
[System.io.path]::GetFullPath("$Env:USERPROFILE\Desktop\Test.docx")

#>