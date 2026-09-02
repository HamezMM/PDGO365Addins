# CLAUDE.md — PDGO365Addins

Build guidance for Claude Code. This repo builds **VSTO add-ins** (Visual Studio Tools
for Office) — managed C#/.NET COM add-ins that load in-process with **desktop Office on
Windows**, primarily **Excel**. They run with full trust: the whole Office object model,
unrestricted file-system access, P/Invoke, any .NET library.

This is **not** the Office.js / "Office Add-ins" web model. Ignore
`learn.microsoft.com/office/dev/add-ins/` — the relevant docs are
`learn.microsoft.com/visualstudio/vsto/`.

## 1. Hard constraints

- **Windows desktop only.** No Mac, no Excel on the web, no iPad.
- **.NET Framework 4.8** (`TargetFrameworkVersion` `v4.8`). VSTO does not run on
  .NET 5+/.NET Core. Do not try to migrate projects to SDK-style / `net8.0`.
- **C#**, language version `9.0` (set in the `.csproj`). WinForms and WPF are available
  for dialogs.
- Every project is **AnyCPU**; the add-in loads into whatever bitness Excel is (this
  machine: 64-bit). Never add x86/x64 build configs.

## 2. Prerequisites — READ BEFORE BUILDING

| Requirement | This machine (checked 2026-09-01) |
| --- | --- |
| Visual Studio 2022 Community 17.14 | `C:\Program Files\Microsoft Visual Studio\2022\Community` |
| **VS workload "Office/SharePoint development"** (`Microsoft.VisualStudio.Workload.Office`) | ✅ installed 2026-09-01 (`--add Microsoft.VisualStudio.Workload.Office --includeRecommended`). See §3 if it ever needs reinstalling. |
| .NET Framework 4.8 targeting pack | ✅ |
| VSTO Runtime v4 | ✅ |
| **Code-signing cert for manifest signing** | ⚠️ **Per-machine, not in repo.** VSTO signs `.vsto`/`.dll.manifest` on every build (plain `msbuild` and F5 both). Create a self-signed dev cert once and put its thumbprint in each project's `.csproj` `<ManifestCertificateThumbprint>` — see `SheetToTxt/SETUP.md` "Signing certificate". |
| Excel | Microsoft 365, 64-bit (`Office16`) |
| MSBuild | `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe` |

## 3. Install the Office workload (one-time, required)

The workload supplies `Microsoft.VisualStudio.Tools.Office.targets` (the VSTO build
targets), the design-time assemblies the projects reference, and the project templates.
`SheetToTxt.csproj` emits a clear build error if the targets are missing.

**Claude may run this install directly.** Prefer the CLI over asking the user. Run from
the Bash or PowerShell tool:

```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\Installer\setup.exe" modify `
  --installPath "C:\Program Files\Microsoft Visual Studio\2022\Community" `
  --add Microsoft.VisualStudio.Workload.Office --includeRecommended `
  --quiet --norestart
```

- Flags are `--quiet --norestart` only. **Do not pass `--wait`** — this installer
  (`setup.exe` 4.4.x) rejects it with `Option 'wait' is unknown` and exit code `87`.
  `setup.exe --quiet` already blocks until the install completes and returns its exit
  code; the install can take 5–15 min, so raise the tool `timeout` to ~900000 ms (or run
  it backgrounded and poll `vswhere`). Exit codes: `0` = success, `3010` = success +
  reboot pending, `1` / other = failure (inspect
  `%TEMP%\dd_setup_*.log` and
  `%ProgramData%\Microsoft\VisualStudio\Packages\_Instances\<id>\logs\`).
- **Elevation:** the installer needs admin rights, and a `--quiet` run **cannot** raise a
  UAC prompt — if the session is not elevated it fails immediately. Check first:
  ```powershell
  ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  ```
  If that is `False`, or the install exits with an elevation error (`5007` =
  "commands with --quiet or --passive should be run elevated from the beginning", `740`,
  `1602`, `1223`), **fall back to asking the user** to run the command from an elevated
  PowerShell, or VS Installer GUI ▸ Modify ▸ check **Office/SharePoint development** ▸
  Modify. On this machine the interactive account (`AzureAD\JamesMckinnon`) is **not** a
  local administrator, so in practice the user has to run it elevated themselves — tell
  them to paste the block above into an elevated shell, or use `! <command>` after
  launching Claude Code elevated.
- Verify afterward:
  `& "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe" -products * -requires Microsoft.VisualStudio.Workload.Office -property installationPath`
  should print the VS path. Also confirm
  `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets`
  now exists.

VS 2026+ deprecates the VSTO templates (the runtime stays supported). If the templates
are gone, copy an existing project folder as the starting point (§6) instead of
File ▸ New.

## 4. Anatomy of an add-in project

`SheetToTxt/` is the reference implementation. Every add-in follows this shape:

| File | Hand-maintained? | Role |
| --- | --- | --- |
| `<Name>.csproj` | ✅ yes | Classic (non-SDK) MSBuild project. `ProjectTypeGuids` = Office flavor + C#. Imports `Microsoft.VisualStudio.Tools.Office.targets`. `ProjectExtensions/FlavorProperties` tells VS it's a VSTO Excel project. |
| `ThisAddIn.cs` | ✅ yes | The `ThisAddIn` partial class you edit: `ThisAddIn_Startup` / `_Shutdown`, and the `CreateRibbonExtensibilityObject` override that returns the ribbon. Also exposes `ThisAddIn.Instance` so feature code reaches `Application` without `Globals`. |
| `ThisAddIn.Designer.cs` | ❌ generated | VSTO plumbing: base class `Microsoft.Office.Tools.AddInBase`, the `Application` field, `Globals`. The VSTO designer regenerates it from the `.xml` on build. Do not edit. |
| `ThisAddIn.Designer.xml` | ❌ generated | "Blueprint" the designer reads to regenerate the `.cs`. Do not edit. |
| `Ribbon/<Name>Ribbon.xml` | ✅ yes | `customUI` markup (namespace `http://schemas.microsoft.com/office/2006/01/customui`). **Build Action = Embedded Resource**, with an explicit `LogicalName` matching the string passed to `GetManifestResourceStream`. |
| `Ribbon/<Name>Ribbon.cs` | ✅ yes | `[ComVisible(true)]` class implementing `Microsoft.Office.Core.IRibbonExtensibility`. `GetCustomUI` returns the embedded XML; `OnLoad` caches the `IRibbonUI`; one thin callback per control. |
| `<Feature>.cs` | ✅ yes | The actual work (e.g. `SheetExporter`, `WorkbookLocation`). Ribbon callbacks stay thin — gather context, call the feature class, show the result/error. |
| `Properties/AssemblyInfo.cs` | ✅ yes | `[assembly: Guid(...)]` is the add-in identity — unique per project. Bump `AssemblyVersion` on release. |
| `app.config` | ✅ yes | `supportedRuntime v4.0 sku=".NETFramework,Version=v4.8"`. |
| `packages.config` | ✅ yes | Empty by default — Office/VSTO assemblies come from the GAC via the build targets, not NuGet. |

Repo root: `PDGO365Addins.sln` (one solution, one project per add-in), `.gitignore`
(bin/obj, `.vs/`, `*.vsto`, `*.manifest`, `*.pfx`), `README.md`, this file.

## 5. Build, run, debug

```
# from repo root, after the workload is installed AND a signing cert exists (SETUP.md)
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  PDGO365Addins.sln /t:Restore,Build /p:Configuration=Debug
```

Build error `The "ManageCertificateStore" task was not given a value for the required
parameter "CertificateThumbprint"` or `the ClickOnce manifest signing option is not
selected` = no signing cert on this machine. Fix per `SheetToTxt/SETUP.md`.

- **F5 in Visual Studio** is the normal loop: builds, writes the COM-add-in registry
  keys under `HKCU\Software\Microsoft\Office\Excel\Addins\<Name>` (`LoadBehavior=3`,
  `Manifest=…\<Name>.vsto|vstolocal`), launches Excel, attaches the managed debugger.
- **A plain build already registers** the add-in for the current user. To unregister:
  Build ▸ **Clean Solution** (or `msbuild /t:Clean`).
- If Excel greys the add-in out after an unhandled exception: File ▸ Options ▸ Add-ins ▸
  Manage **Disabled Items** and **COM Add-ins**, re-enable, restart Excel.
- Add-in fails to load silently: check Windows Event Viewer ▸ Application, and confirm
  `bin\Debug\<Name>.vsto` and `<Name>.dll.manifest` exist.

## 6. Recipe — add a new add-in to the repo

1. Copy `SheetToTxt/` to `<NewName>/`. Rename `SheetToTxt.csproj` → `<NewName>.csproj`.
2. In the `.csproj`: new `<ProjectGuid>` (fresh GUID), set `RootNamespace` /
   `AssemblyName` / `GeneratedCodeNamespace` (in `ProjectExtensions`) to `<NewName>`.
3. In `AssemblyInfo.cs`: new `[assembly: Guid]`, update title/description.
4. Rename the `namespace` in every `.cs`/`.xml` from `SheetToTxt` to `<NewName>`
   (including `ThisAddIn.Designer.xml` `hostitem:namespace` and the ribbon
   `LogicalName` / `GetManifestResourceStream` string).
5. For a **non-Excel host** (Word/PowerPoint/Outlook): change `<OfficeApplication>`, the
   `ProjectExtensions` `ApplicationType`/`HostName`/`DebugInfoExeName`, the interop
   reference (`Microsoft.Office.Interop.Word` etc.), the designer's `factoryType` /
   `hostObject` type, and the ribbon `tab idMso`. Easiest is to generate that project
   from the VS template once the workload is installed and port the feature classes in.
6. `dotnet sln PDGO365Addins.sln add <NewName>\<NewName>.csproj` — or add it in VS —
   then add the same `{ProjectConfigurationPlatforms}` lines as the existing project.

## 7. Recipe — add a feature + ribbon button

1. **Feature class**: `MyFeature.cs`, one public entry method that returns a small result
   type or throws with a user-readable `Message`. Reach Excel via
   `ThisAddIn.Instance.Application`. Follow the interop rules in §8.
2. **Ribbon XML**: add a `<button>` inside the group in `Ribbon/<Name>Ribbon.xml`:
   ```xml
   <button id="MyFeatureButton" label="Do The Thing" size="large"
           imageMso="FileSaveAs"
           onAction="MyFeatureButton_OnAction"
           getEnabled="MyFeatureButton_GetEnabled"
           screentip="…" supertip="…" />
   ```
3. **Callbacks** in `Ribbon/<Name>Ribbon.cs` — names must match the XML exactly, must be
   `public`, and the signatures must be exact (a wrong signature compiles but silently
   does nothing — there is no IntelliSense for these):
   - `onAction` (button): `public void Name(Office.IRibbonControl control)`
   - `getEnabled`: `public bool Name(Office.IRibbonControl control)`
   - `getImage`: `public stdole.IPictureDisp Name(Office.IRibbonControl control)`
   - `onLoad` (customUI): `public void OnLoad(Office.IRibbonUI ribbon)`
   Call `_ribbon.InvalidateControl("MyFeatureButton")` after state changes so
   `getEnabled` re-runs.
4. Wrap the callback body in try/catch (see §9).

## 8. Office interop rules

- **Main (STA) thread only.** Never call the object model from a `Task`, `Timer`, or
  thread-pool callback. If you must do async work, capture `SynchronizationContext` on
  startup and post back.
- **No two-dot chains** on COM objects — `wb.Worksheets[1].Range["A1"]` leaks the
  intermediate RCW. Assign each step to a local.
- **Bulk cell access**: one `Range.Value2` get/set of an `object[,]` (1-based bounds; a
  single cell comes back as a scalar, not an array). Never loop cell-by-cell. `Value2`
  gives raw values; `.Text` gives the displayed string but is per-cell and slow.
- To emit a sheet as delimited text, prefer letting Excel do it —
  `Worksheet.Copy()` into a throwaway workbook then `Workbook.SaveAs(path,
  XlFileFormat.xlTextWindows)` — instead of formatting values by hand. That's what
  `SheetExporter` does.
- Bracket document mutations with `Application.ScreenUpdating = false` and, if it
  recalculates, `Application.Calculation = xlCalculationManual`; restore both in
  `finally`. Same for `Application.DisplayAlerts` when calling `SaveAs`/`Close`.
- `Marshal.ReleaseComObject` matters for long loops and burst-allocated objects; for a
  one-shot ribbon command, GC at shutdown is acceptable. Be consistent within a feature.
- Useful workbook members: `ActiveWorkbook.Path` (folder; `""` until first save),
  `.FullName`, `.Name`; `ActiveSheet` cast to `Excel.Worksheet` (may be a chart sheet —
  check).

## 9. Error handling & UX

- **No exception may escape a ribbon callback** — Office can hard-disable the add-in.
  Every callback: `try { … } catch (Exception ex) { MessageBox.Show(ex.Message, "<Add-in>",
  MessageBoxButtons.OK, MessageBoxIcon.Warning); }`.
- Throw `InvalidOperationException` with a plain-language message for expected failure
  states (no workbook open, workbook unsaved, wrong sheet type). The callback shows
  `ex.Message` directly, so write it for the end user.
- `getEnabled` should catch internally and return `false` rather than throw.
- Keep `ThisAddIn_Startup` fast — no I/O, no dialogs. Slow startup gets the add-in
  disabled by Office's performance monitor.

## 10. Conventions

- One feature = one class with a clear entry point. Ribbon and `ThisAddIn` stay thin.
- Interop references use `EmbedInteropTypes = true` (type embedding) so an add-in isn't
  pinned to one Office version's PIA.
- Don't hand-edit `*.Designer.*`.
- No secrets in source. Per-user config under `%APPDATA%\PDG\<AddInName>\`.
- Match the C# style of the PDG Revit add-in repos (same team) —
  see `..\PDGRevitAddinCollection`.
- Every manifest/behavior change: bump `AssemblyVersion` and note it in the add-in's
  `SETUP.md` / changelog.

## 11. Deployment

VSTO add-ins are activated by registry keys pointing Office at a `.vsto` deployment
manifest.

| Method | Use |
| --- | --- |
| **F5 / local build** | Dev only; registers for the current user. |
| **ClickOnce** (VS ▸ Publish → network share) | Internal rollout to PDG staff; auto-updates on version bump. Manifests must be signed — VS makes a test cert; use a real code-signing cert for production. Set `SignManifests=true` + `ManifestCertificateThumbprint` in the `.csproj` for this. |
| **MSI (WiX / Advanced Installer)** | When IT wants a per-machine managed package. The installer writes `HKLM\…\Office\Excel\Addins\<Name>` with `FriendlyName`, `Description`, `LoadBehavior=3`, `Manifest=<path>\<Name>.vsto|vstolocal`. |

End-user machines need the **VSTO Runtime** (ships with modern Office) and **.NET
Framework 4.8** (in-box on Windows 10 1903+ / 11).

## 12. References

- VSTO docs root: <https://learn.microsoft.com/visualstudio/vsto/>
- First Excel VSTO add-in: <https://learn.microsoft.com/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-excel>
- Program VSTO add-ins (ThisAddIn / Globals): <https://learn.microsoft.com/visualstudio/vsto/programming-vsto-add-ins>
- Ribbon XML + callbacks: <https://learn.microsoft.com/visualstudio/vsto/ribbon-xml>
- customUI element/attribute reference: <https://learn.microsoft.com/previous-versions/office/developer/office-2007/aa338199(v=office.12)>
- Excel interop API: <https://learn.microsoft.com/dotnet/api/microsoft.office.interop.excel>
- Deploy an Office solution: <https://learn.microsoft.com/visualstudio/vsto/deploying-an-office-solution>
- Release COM objects: <https://learn.microsoft.com/visualstudio/vsto/systematically-releasing-objects>
