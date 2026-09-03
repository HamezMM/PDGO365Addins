<#
.SYNOPSIS
    Builds SheetToTxt in Release and publishes the ClickOnce deployment into the
    synced SharePoint folder the team installs from.

.DESCRIPTION
    Runs `msbuild /t:Publish` (Release), which produces setup.exe, SheetToTxt.vsto and
    Application Files\ under PublishDir. Also drops PDG-CodeSigning.cer and
    Install-SheetToTxt.ps1 next to it so a team member has everything in one folder.

    Prerequisites on the publishing machine:
      - VS 2022 with the Office/SharePoint workload (see ..\CLAUDE.md section 3).
      - The "PDG Code Signing" cert WITH private key in Cert:\CurrentUser\My
        (import PDG-CodeSigning.pfx once - see README.md).
      - The SharePoint library "Peake Design - Documents" synced locally.

.PARAMETER PublishDir
    Local path of the synced target folder. Default is the standard mount point for
    "...\Peake Design - Documents\SOFTWARE RESOURCES\O365\SheetToTxt".

.PARAMETER Version
    Four-part ClickOnce version, e.g. 1.0.1.0. Defaults to bumping the last part of
    whatever is in the .csproj.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Publish-SheetToTxt.ps1 -Version 1.1.0.0
#>
[CmdletBinding()]
param(
    [string]$PublishDir = (Join-Path $env:USERPROFILE 'Peake Design Group\Peake Design - Documents\SOFTWARE RESOURCES\O365\SheetToTxt'),
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$repo    = Split-Path $PSScriptRoot -Parent
$csproj  = Join-Path $repo 'SheetToTxt\SheetToTxt.csproj'
$msbuild = 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe'

# --- version -----------------------------------------------------------------
if (-not $Version) {
    $current = ([xml](Get-Content $csproj)).Project.PropertyGroup.ApplicationVersion |
        Where-Object { $_ } | Select-Object -First 1
    $p = $current.Split('.'); $p[3] = [int]$p[3] + 1
    $Version = ($p -join '.')
}
Write-Host "Publishing version $Version -> $PublishDir"

# --- signing cert present? -------------------------------------------------
$thumb = 'CD2816561ED07ACF6C0F05C705A43B5197AB7F32'
if (-not (Test-Path "Cert:\CurrentUser\My\$thumb")) {
    throw "PDG Code Signing cert ($thumb) not in Cert:\CurrentUser\My. Import PDG-CodeSigning.pfx first - see README.md."
}

New-Item -ItemType Directory -Force -Path $PublishDir | Out-Null

# --- publish ---------------------------------------------------------------
# Pass args via an MSBuild response file: PublishDir contains spaces and a
# trailing separator, which PowerShell's native-arg quoting mangles ("\"" eats
# the following tokens). A .rsp file is read literally - no shell parsing.
$rsp = [System.IO.Path]::GetTempFileName() + '.rsp'
@(
    "`"$csproj`""
    '/t:Publish'
    '/p:Configuration=Release'
    "/p:ApplicationVersion=$Version"
    # trailing '\\' inside the quotes -> one '\' after arg parsing (a lone '\"'
    # would escape the quote and swallow the next tokens).
    "/p:PublishDir=`"$($PublishDir.TrimEnd('\'))\\`""
    '/v:minimal'
    '/nologo'
) | Set-Content -Path $rsp -Encoding UTF8
try {
    & $msbuild "@$rsp"
    if ($LASTEXITCODE -ne 0) { throw "msbuild /t:Publish failed ($LASTEXITCODE)" }
}
finally {
    Remove-Item $rsp -ErrorAction SilentlyContinue
}

# --- drop the install helpers alongside ----------------------------------
Copy-Item (Join-Path $PSScriptRoot 'PDG-CodeSigning.cer')      $PublishDir -Force
Copy-Item (Join-Path $PSScriptRoot 'Install-SheetToTxt.ps1')   $PublishDir -Force
Copy-Item (Join-Path $PSScriptRoot 'Uninstall-SheetToTxt.ps1') $PublishDir -Force
Copy-Item (Join-Path $PSScriptRoot 'INSTALL.md')               $PublishDir -Force

Write-Host "`nPublished. Contents of $PublishDir :"
Get-ChildItem $PublishDir | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize
Write-Host @"

Next:
  1. Let OneDrive finish syncing $PublishDir to SharePoint.
  2. Bump AssemblyVersion in SheetToTxt\Properties\AssemblyInfo.cs and the
     <ApplicationVersion> in the .csproj to $Version, then commit.
  3. Team members run Install-SheetToTxt.ps1 from their synced copy of that folder.
"@
