# Registers the Shredder Mail Add-in for the current user (no admin required).
# Run from the folder containing OutlookShredder.MailAddin.dll after publishing.

$ErrorActionPreference = 'Stop'

$dllPath = Resolve-Path (Join-Path $PSScriptRoot 'OutlookShredder.MailAddin.dll')
$classGuid    = '{7F3A1D5C-4E2B-4A8F-9C6D-1B3E5F7A9C2D}'
$progId       = 'OutlookShredder.MailAddin.Connect'
$className    = 'OutlookShredder.MailAddin.Connect'
$assemblyName = 'OutlookShredder.MailAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$codeBase     = 'file:///' + $dllPath.ToString().Replace('\', '/')
# .NET Framework component category GUID (required for managed COM objects)
$dotNetCatGuid = '{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'

Write-Host "DLL: $dllPath"
Write-Host "Registering COM class in HKCU (no admin required)..."

# CLSID registration
$clsidRoot = "HKCU:\Software\Classes\CLSID\$classGuid"
New-Item -Path $clsidRoot -Force | Out-Null
Set-ItemProperty -Path $clsidRoot -Name '(Default)' -Value $className

$inproc = "$clsidRoot\InprocServer32"
New-Item -Path $inproc -Force | Out-Null
Set-ItemProperty -Path $inproc -Name '(Default)'      -Value 'mscoree.dll'
Set-ItemProperty -Path $inproc -Name 'Class'          -Value $className
Set-ItemProperty -Path $inproc -Name 'Assembly'       -Value $assemblyName
Set-ItemProperty -Path $inproc -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path $inproc -Name 'ThreadingModel' -Value 'Both'
Set-ItemProperty -Path $inproc -Name 'CodeBase'       -Value $codeBase

# Version-specific sub-key (required by regasm pattern)
$inprocVer = "$inproc\1.0.0.0"
New-Item -Path $inprocVer -Force | Out-Null
Set-ItemProperty -Path $inprocVer -Name 'Class'          -Value $className
Set-ItemProperty -Path $inprocVer -Name 'Assembly'       -Value $assemblyName
Set-ItemProperty -Path $inprocVer -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path $inprocVer -Name 'CodeBase'       -Value $codeBase

$progIdKey = "$clsidRoot\ProgId"
New-Item -Path $progIdKey -Force | Out-Null
Set-ItemProperty -Path $progIdKey -Name '(Default)' -Value $progId

# Implemented Categories - marks this as a .NET Framework managed component
New-Item -Path "$clsidRoot\Implemented Categories\$dotNetCatGuid" -Force | Out-Null

# ProgId -> CLSID mapping
$progIdRoot = "HKCU:\Software\Classes\$progId"
New-Item -Path "$progIdRoot" -Force | Out-Null
Set-ItemProperty -Path "$progIdRoot" -Name '(Default)' -Value $className
New-Item -Path "$progIdRoot\CLSID" -Force | Out-Null
Set-ItemProperty -Path "$progIdRoot\CLSID" -Name '(Default)' -Value $classGuid

Write-Host "Registering Outlook add-in..."

$addinKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name 'Description'  -Value 'Shredder Mail Add-in'
Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'Shredder Mail Add-in'
Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Type DWord -Value 3

Write-Host "Done. Close and restart Outlook to activate the add-in."
