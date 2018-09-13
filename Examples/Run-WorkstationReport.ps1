Import-Module PSWriteWord
Import-Module PSSharedGoods
Import-Module PSWinDocumentation

$FilePath = "$Env:USERPROFILE\Desktop\PSWinDocumentation-WorkstationReport.docx"

Start-WinDocumentationWorkstation -ComputerName 'EVO1' -FilePath $FilePath -OpenDocument