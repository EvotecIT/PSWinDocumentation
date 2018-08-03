Import-Module PSWInDocumentation -Force
Import-Module PSWriteWord -Force
Import-Module C:\Support\GitHub\ImportExcel\ImportExcel.psd1 -Force
Import-Module ActiveDirectory

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.docx"
$FilePathExcel = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Report.xlsx"

Clear-Host
Start-ActiveDirectoryDocumentation -CompanyName 'Euvic' -FilePath $FilePath -Verbose -FilePathExcel $FilePathExcel #-OpenDocument