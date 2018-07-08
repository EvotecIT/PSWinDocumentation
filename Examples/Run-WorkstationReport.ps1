Import-Module PSWriteWord #-Force
Import-Module PSWinDocumentation -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWinDocumentation-WorkstationReport.docx"

Start-WinDocumentationWorkstation -ComputerName 'Evo1' -FilePath $FilePath