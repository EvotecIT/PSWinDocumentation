Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord -Force
Import-Module ActiveDirectory

$FilePathTemplate = "$PSScriptRoot\Templates\WordTemplate.docx"
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"

Clear-Host
Start-ActiveDirectoryDocumentation -CompanyName 'Evotec' -FilePathTemplate $FilePathTemplate -FilePath $FilePath -OpenDocument