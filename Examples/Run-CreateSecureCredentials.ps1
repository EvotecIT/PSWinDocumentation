$MyPassword = ""
$FileToSaveSecurePassword = 'C:\Support\GitHub\PSWinDocumentation\Ignore\MySecurePassword.txt'

# Option 1
$SecurePassword = $MyPassword | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
$SecurePassword | Out-File $FileToSaveSecurePassword


# Option 2
#(Get-Credential).Password | ConvertFrom-SecureString | Out-File $FileToSaveSecurePassword


# Option 3
#Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $FileToSaveSecurePassword


# Read Secure Password
$SecurePasswordFromFile = Get-Content $FileToSaveSecurePassword | ConvertTo-SecureString
$SecurePasswordFromFile