<#
.SYNOPSIS
    Removes the SheetToTxt Excel add-in for the current user.

.DESCRIPTION
    Uninstalls the ClickOnce deployment (via the Windows "Installed apps" entry) and,
    with -RemoveCert, also drops the PDG Code Signing certificate from the current
    user's trust stores. Leave the cert in place if other PDG add-ins use it.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Uninstall-SheetToTxt.ps1
#>
[CmdletBinding()]
param(
    [switch]$RemoveCert
)

$ErrorActionPreference = 'Continue'

# --- ClickOnce uninstall -----------------------------------------------------
$uninstall = Get-ChildItem 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall' |
    ForEach-Object { Get-ItemProperty $_.PSPath } |
    Where-Object { $_.DisplayName -like 'SheetToTxt*' }

if ($uninstall) {
    Write-Host "Uninstalling: $($uninstall.DisplayName)"
    # ClickOnce UninstallString is a rundll32 call; run it (it shows a confirm dialog).
    Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', $uninstall.UninstallString -Wait
} else {
    Write-Host "No SheetToTxt ClickOnce install found for this user."
}

# Clear the COM add-in registration if a plain build left one behind.
$addinKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\SheetToTxt'
if (Test-Path $addinKey) { Remove-Item $addinKey -Force; Write-Host "Removed $addinKey" }

# --- optional: untrust the signing cert ------------------------------------
if ($RemoveCert) {
    $thumb = 'CD2816561ED07ACF6C0F05C705A43B5197AB7F32'
    foreach ($storeName in 'Root', 'TrustedPublisher') {
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new(
            $storeName, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        $store.Open('ReadWrite')
        foreach ($c in $store.Certificates.Find('FindByThumbprint', $thumb, $false)) {
            $store.Remove($c); Write-Host "Removed cert from CurrentUser\$storeName"
        }
        $store.Close()
    }
}

Write-Host "`nDone. Restart Excel."
