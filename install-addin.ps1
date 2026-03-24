# install-addin.ps1
# Registers the OutlookShredder manifest as a trusted Outlook add-in catalog.
# Run once as the current user (no admin required).
# After running: restart Outlook -> Get Add-ins -> My Add-ins -> OutlookShredder section.

$ManifestDir  = "$PSScriptRoot\OutlookShredder.AddinHost"
$ManifestFile = "$ManifestDir\manifest.xml"

if (-not (Test-Path $ManifestFile)) {
    Write-Error "manifest.xml not found at $ManifestFile"
    exit 1
}

# Outlook WEF trusted catalogs registry path
$CatalogsRoot = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"

# Use a stable GUID so re-running is idempotent
$CatalogGuid  = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
$CatalogKey   = "$CatalogsRoot\$CatalogGuid"

if (-not (Test-Path $CatalogsRoot)) {
    New-Item -Path $CatalogsRoot -Force | Out-Null
}

New-Item      -Path $CatalogKey -Force | Out-Null
Set-ItemProperty -Path $CatalogKey -Name "Id"          -Value $CatalogGuid
Set-ItemProperty -Path $CatalogKey -Name "Url"         -Value $ManifestDir
Set-ItemProperty -Path $CatalogKey -Name "Flags"       -Value 1 -Type DWord   # enabled
Set-ItemProperty -Path $CatalogKey -Name "ShowInMenu"  -Value 1 -Type DWord   # show in My Add-ins

Write-Host ""
Write-Host "SUCCESS: Catalog registered." -ForegroundColor Green
Write-Host ""
Write-Host "Catalog folder : $ManifestDir"
Write-Host "Registry key   : $CatalogKey"
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Make sure OutlookShredder.AddinHost is running (https://localhost:3000)"
Write-Host "  2. Close and reopen Outlook"
Write-Host "  3. Open any email -> Home ribbon -> Get Add-ins -> My Add-ins"
Write-Host "  4. Scroll to the 'OutlookShredder.AddinHost' section -> click Add"
Write-Host ""
