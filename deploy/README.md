# deploy/ — sharing SheetToTxt with the team

SheetToTxt is a **VSTO add-in**. Office only loads it if its deployment manifests are
signed by a certificate the machine **trusts**, so distribution has two parts:

1. the signed ClickOnce build, published to a shared folder, and
2. the signing certificate's **public** half, trusted on each user's machine.

Nobody but the person who publishes needs the private key.

| File | In git? | What it is |
| --- | --- | --- |
| `PDG-CodeSigning.cer` | ✅ yes | Public cert. Safe to share. Team machines trust this. Thumbprint `CD2816561ED07ACF6C0F05C705A43B5197AB7F32`, expires 2031-09-02. |
| `PDG-CodeSigning.pfx` | ❌ **never** | Cert **+ private key**. Currently at `%USERPROFILE%\Documents\PDG-Signing\` on the publisher's machine, password in `password.txt` beside it. **Move the password into a shared password manager and delete `password.txt`; back the `.pfx` up there too.** If the key is lost, a new cert means every team machine must re-trust. |
| `Install-SheetToTxt.ps1` | ✅ yes | Team members run this: trusts the cert, runs the ClickOnce installer. |
| `Uninstall-SheetToTxt.ps1` | ✅ yes | Removes the add-in (`-RemoveCert` also untrusts the cert). |
| `Publish-SheetToTxt.ps1` | ✅ yes | Publisher runs this: Release build → ClickOnce publish into the synced SharePoint folder. |

Publish target (OneDrive-synced SharePoint):
`…\Peake Design - Documents\SOFTWARE RESOURCES\O365\SheetToTxt`
= `https://chesapeakeud.sharepoint.com/sites/Peake/Shared Documents/SOFTWARE RESOURCES/O365/SheetToTxt/`

---

## One-time: set up the publishing machine

1. Install the VS Office/SharePoint workload if needed (`../CLAUDE.md` §3).
2. Import the signing cert **with its private key** into your personal store:
   ```powershell
   $pw = Get-Content "$env:USERPROFILE\Documents\PDG-Signing\password.txt"
   certutil -f -user -p $pw -importpfx My "$env:USERPROFILE\Documents\PDG-Signing\PDG-CodeSigning.pfx" NoRoot
   ```
   (Another maintainer: get the `.pfx` + password from whoever holds it, over a secure
   channel — not email, not the repo.)
3. Make sure the `Peake Design - Documents` library is synced locally.

## Cut a release

```powershell
cd deploy
powershell -ExecutionPolicy Bypass -File .\Publish-SheetToTxt.ps1 -Version 1.1.0.0
```

Then:
- bump `AssemblyVersion` in `SheetToTxt/Properties/AssemblyInfo.cs` and
  `<ApplicationVersion>` in `SheetToTxt/SheetToTxt.csproj` to match, and commit;
- wait for OneDrive to finish uploading the folder.

Existing installs pick the new version up automatically on next Excel launch
(`UpdateMode=Foreground`).

## Team install (self-service)

Send people this:

> 1. In File Explorer open
>    **Peake Design Group ▸ Peake Design - Documents ▸ SOFTWARE RESOURCES ▸ O365 ▸ SheetToTxt**
>    (the SharePoint folder — it should be synced; if not, open it in the browser and
>    "Add shortcut to OneDrive").
> 2. Right-click **Install-SheetToTxt.ps1 ▸ Run with PowerShell**.
>    (If it's blocked: open PowerShell, run
>    `powershell -ExecutionPolicy Bypass -File "<that path>\Install-SheetToTxt.ps1"`.)
> 3. Click **Install** on the "Customization Installer" prompt.
> 4. Restart Excel. The **Home ▸ Export Sheet to .txt** button appears.

To remove: run `Uninstall-SheetToTxt.ps1` from the same folder.

## If SharePoint won't serve the ClickOnce update

Some tenants block or auth-gate direct file GETs. Symptoms: update check hangs or
prompts for sign-in. Options:
- Team installs/updates by running `setup.exe` from their **synced local copy** instead
  of the URL (`Install-SheetToTxt.ps1` already does this).
- Or move the publish target to a plain file share (`\\server\share\...`) and set
  `InstallUrl` / `UpdateUrl` in the `.csproj` to that UNC path.

## Renewing the cert (before 2031, or if the key leaks)

1. `New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=PDG Code Signing, O=Peake Design, C=CA' -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable -KeyLength 3072 -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(5)`
2. Export new `.cer` (commit) and `.pfx` (secure store).
3. Update `<ManifestCertificateThumbprint>` in the `.csproj`, the thumbprints in
   `Install-SheetToTxt.ps1` / `Uninstall-SheetToTxt.ps1`, and the table above.
4. Publish a new version. Team re-runs `Install-SheetToTxt.ps1` (it re-trusts).
