$classGuid = '{7F3A1D5C-4E2B-4A8F-9C6D-1B3E5F7A9C2D}'
$progId    = 'OutlookShredder.MailAddin.Connect'

Remove-Item "HKCU:\Software\Classes\CLSID\$classGuid"     -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Classes\$progId"               -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Unregistered. Restart Outlook to complete removal."
