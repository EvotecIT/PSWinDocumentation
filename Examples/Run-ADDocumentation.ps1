Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord -Force
Import-Module C:\Support\GitHub\ImportExcel\ImportExcel.psd1 -Force # Modified version of original
Import-Module ActiveDirectory

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"
$FilePathExcel = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.xlsx"

Clear-Host
Start-ActiveDirectoryDocumentation -CompanyName 'Evotec' -FilePath $FilePath -Verbose -FilePathExcel $FilePathExcel -OpenDocument