<#
.SYNOPSIS
    Installs the SheetToTxt Excel add-in for the current user (no admin rights).

.DESCRIPTION
    1. Trusts the "PDG Code Signing" certificate: adds PDG-CodeSigning.cer to the
       current user's Trusted Root and Trusted Publishers stores, so Office loads the
       signed VSTO manifest without a "publisher cannot be verified" prompt.
    2. Runs the ClickOnce installer so Excel picks the add-in up next launch.

    Idempotent - safe to re-run (e.g. after the cert is renewed).

.PARAMETER Source
    Folder holding PDG-CodeSigning.cer and the published ClickOnce output
    (setup.exe / SheetToTxt.vsto). Defaults to the folder this script sits in.
    Normal use: open the synced SharePoint folder
    "...\Peake Design - Documents\SOFTWARE RESOURCES\O365\SheetToTxt" and run this.

.PARAMETER FromUrl
    Install from the published https URL instead of a local setup.exe. Use if the
    synced copy isn't available.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Install-SheetToTxt.ps1
#>
[CmdletBinding()]
param(
    [string]$Source = $PSScriptRoot,
    [switch]$FromUrl
)

$ErrorActionPreference = 'Stop'
$InstallUrl = 'https://chesapeakeud.sharepoint.com/sites/Peake/Shared%20Documents/SOFTWARE%20RESOURCES/O365/SheetToTxt/'

# --- 1. Trust the signing certificate (current user) --------------------------
$cer = Join-Path $Source 'PDG-CodeSigning.cer'
if (-not (Test-Path $cer)) { throw "PDG-CodeSigning.cer not found in '$Source'." }

$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($cer)
Write-Host "Signing cert : $($cert.Subject)"
Write-Host "Thumbprint   : $($cert.Thumbprint)"
Write-Host "Expires      : $($cert.NotAfter.ToString('yyyy-MM-dd'))"

foreach ($storeName in 'Root', 'TrustedPublisher') {
    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new(
        $storeName, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $store.Open('ReadWrite')
    $have = $store.Certificates.Find('FindByThumbprint', $cert.Thumbprint, $false).Count -gt 0
    if (-not $have) { $store.Add($cert); Write-Host "Added to CurrentUser\$storeName" }
    else            { Write-Host "Already trusted in CurrentUser\$storeName" }
    $store.Close()
}

# --- 2. Kick off the ClickOnce install ---------------------------------------
if ($FromUrl) {
    Write-Host "`nOpening $InstallUrl ..."
    Start-Process "$InstallUrl/setup.exe"
}
else {
    $installer = 'setup.exe', 'SheetToTxt.vsto' |
        ForEach-Object { Join-Path $Source $_ } |
        Where-Object   { Test-Path $_ } |
        Select-Object  -First 1
    if (-not $installer) {
        Write-Warning "No setup.exe / SheetToTxt.vsto in '$Source'. Cert is trusted; re-run with -FromUrl, or point -Source at the published folder."
        return
    }
    Write-Host "`nLaunching $installer ..."
    Start-Process -FilePath $installer -Wait
}

Write-Host "`nDone. Close and reopen Excel - the 'Home > Export Sheet to .txt' button should appear."
