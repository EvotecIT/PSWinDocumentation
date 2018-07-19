Import-Module PSWriteWord #-Force
Import-Module PSWinDocumentation -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWinDocumentation-WorkstationReport.docx"

Start-WinDocumentationWorkstation -ComputerName 'EVO1' -FilePath $FilePath -OpenDocument