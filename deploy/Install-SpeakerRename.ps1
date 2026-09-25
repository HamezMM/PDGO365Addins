<#
.SYNOPSIS
    Installs the SpeakerRename Word add-in for the current user (no admin rights).

.DESCRIPTION
    1. Trusts the "PDG Code Signing" certificate: adds PDG-CodeSigning.cer to the
       current user's Trusted Root and Trusted Publishers stores, so Office loads the
       signed VSTO manifest without a "publisher cannot be verified" prompt.
    2. Runs Office's own VSTO installer (VSTOInstaller.exe /i) on the LOCAL synced
       SpeakerRename.vsto, so Word picks the add-in up next launch.

    Why not setup.exe: it is a web bootstrapper that downloads from the SharePoint https
    URL, and SharePoint answers 403 Forbidden because the installer can't sign in to
    Microsoft 365. VSTOInstaller on the local .vsto never touches the web. Updates are
    then checked against that same local .vsto, which OneDrive keeps in sync.

    Idempotent - safe to re-run (e.g. after the cert is renewed).

.PARAMETER Source
    Folder holding PDG-CodeSigning.cer and the published SpeakerRename.vsto (+ Application
    Files\). Defaults to the folder this script sits in, falling back to the standard
    synced path "%USERPROFILE%\Peake Design Group\Peake Design - Documents\SOFTWARE
    RESOURCES\O365\SpeakerRename" when that can't be determined (e.g. pasted into a console).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Install-SpeakerRename.ps1
#>
[CmdletBinding()]
param(
    [string]$Source
)

$ErrorActionPreference = 'Stop'
$DefaultSource = Join-Path $env:USERPROFILE 'Peake Design Group\Peake Design - Documents\SOFTWARE RESOURCES\O365\SpeakerRename'

# $PSScriptRoot is blank when the script isn't run from a file (pasted, piped to iex),
# which used to leave $Source empty and fail with "not found in ''".
if (-not $Source) { $Source = $PSScriptRoot }
if (-not $Source -and $MyInvocation.MyCommand.Path) { $Source = Split-Path $MyInvocation.MyCommand.Path -Parent }
if (-not $Source) { $Source = $DefaultSource }
if (-not (Test-Path $Source)) {
    throw "Folder '$Source' not found. Sync the SharePoint folder (see INSTALL.md) or pass -Source <folder with SpeakerRename.vsto>."
}
$Source = (Resolve-Path $Source).ProviderPath
Write-Host "Installing from: $Source"

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

# --- 2. Install from the local .vsto with Office's VSTO installer -------------
$vsto = Join-Path $Source 'SpeakerRename.vsto'
if (-not (Test-Path $vsto)) {
    throw "SpeakerRename.vsto not found in '$Source'. Make sure the SharePoint folder has finished syncing, or pass -Source."
}

$vstoInstaller = ${env:CommonProgramFiles}, ${env:CommonProgramFiles(x86)}, ${env:CommonProgramW6432} |
    Where-Object   { $_ } |
    ForEach-Object { Join-Path $_ 'microsoft shared\VSTO\10.0\VSTOInstaller.exe' } |
    Where-Object   { Test-Path $_ } |
    Select-Object  -First 1
if (-not $vstoInstaller) {
    throw "VSTOInstaller.exe not found. Install the 'Microsoft Visual Studio 2010 Tools for Office Runtime' (free, from Microsoft) and re-run."
}

Write-Host "`nLaunching Office Customization Installer for $vsto ..."
Write-Host "Click 'Install' in the window that opens (it may be behind other windows)."
$p = Start-Process -FilePath $vstoInstaller -ArgumentList '/i', "`"$vsto`"" -Wait -PassThru
if ($p.ExitCode -ne 0) {
    throw "VSTOInstaller exited with code $($p.ExitCode). If you clicked Cancel, re-run; otherwise see INSTALL.md > Troubleshooting."
}

Write-Host "`nDone. Close and reopen Word - the 'Home > Rename Speakers' button should appear."
