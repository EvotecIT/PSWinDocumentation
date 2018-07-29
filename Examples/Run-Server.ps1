Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord # -Force
Import-Module ActiveDirectory

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"

Clear-Host
Start-ActiveDirectoryDocumentation -CompanyName 'Evotec' -FilePath $FilePath -OpenDocument