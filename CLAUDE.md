# CLAUDE.md — PDGO365Addins

Guidance for Claude Code when working in this repository.

## 1. Purpose

This repo holds **custom Office Add-ins for the Microsoft 365 / Office suite** (Excel,
Word, PowerPoint, Outlook, OneNote). Each add-in is a web application that runs inside an
Office client and uses the **Office JavaScript API (Office.js)** to read and write
document content, extend the Office UI, and call external services.

The canonical Microsoft reference is
<https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview>. When a
question is not answered here, consult Microsoft Learn under `/office/dev/add-ins/`
rather than guessing — the platform changes frequently.

## 2. Repository layout

One folder per add-in at the repo root. Do not mix two add-ins in one project folder.

```
PDGO365Addins/
  <addin-name>/            # a single Yo Office project
    manifest.xml           # (or manifest.json for unified manifest)
    package.json
    webpack.config.js
    src/
      taskpane/            # task pane HTML/CSS/TS + Office.onReady wiring
      commands/            # ribbon button / function-command handlers
    assets/                # icons referenced by the manifest (16/32/64/80 px)
  CLAUDE.md
  README.md
```

When creating a new add-in, scaffold it into its own subfolder (see §5).

## 3. Anatomy of an Office Add-in

Every add-in has **two parts**:

1. **Manifest** — an XML or JSON file describing the add-in's identity, the Office apps it
   runs in, permissions, UI integration (ribbon tabs/buttons, context menus), icon URLs,
   and the URL of the web app. The manifest is the unit of sideloading and publishing.
2. **Web application** — HTML/CSS/JS (or TypeScript/React) that renders task panes,
   content add-ins, and dialogs, and calls Office.js. It is an ordinary web app: it can
   also call REST services, do auth, etc. Must be served over **HTTPS**.

The add-in has **no server-side component by default**. Add a middle tier only when you
need SSO token exchange, secrets, or server-only APIs.

## 4. Manifest: which type to use

| Manifest | Format | File | Use when |
| --- | --- | --- | --- |
| **Add-in only manifest** | XML | `manifest.xml` | Default for **production Excel / Word / PowerPoint / OneNote / Project** add-ins today. |
| **Unified manifest for Microsoft 365** | JSON | `manifest.json` | **Outlook** add-ins (recommended), or when the add-in ships alongside other M365 extensions (Teams tab, etc.) as one installable unit. Still **preview** for Excel/Word/PowerPoint — do not use for production in those hosts. |

Rules:

- **Raise the `Version` / `version` on every manifest change.** Deployed add-ins do not
  update for users until the version increases (and admin re-consent is needed for
  permission, scope, or event changes).
- The add-in **`Id` must be a unique GUID**. Generate a fresh one per add-in; never reuse.
- All URLs in the manifest (web app, icons) must be **HTTPS**. Icon-hosting servers must
  not send `Cache-Control: no-store/no-cache`.
- List every external domain the task pane navigates to under `AppDomains`
  (`validDomains` in the unified manifest) or desktop Office opens it in a new browser
  window instead of the pane.
- Validate before every commit and before sideloading:
  ```
  npm run validate            # wraps office-addin-manifest validate
  npx office-addin-manifest validate manifest.xml
  ```

## 5. Creating a new add-in (Yo Office)

Node.js **Active LTS** and npm are prerequisites (this machine: Node 24, npm 11).

```powershell
npm install -g yo generator-office      # once; re-run to update
cd C:\Users\JamesMckinnon\source\repos\PDGO365Addins
yo office                               # run in PowerShell, NOT a bash shell
```

Yo Office prompts, in order:

1. **Project type** — usually `Office Add-in Task Pane project`. Other options: React
   variant, Excel Custom Functions (shared-runtime or JS-only), SSO, Nested App Auth,
   `manifest only` (use when swapping bundler/framework, e.g. Vue).
2. **Language** — prefer **TypeScript**.
3. **Name** — becomes the folder name and the manifest display name.
4. **Office application** — Excel, OneNote, Outlook, PowerPoint, Project, or Word. Pick
   one; broaden the manifest later. **Outlook cannot be combined** with any other host.
5. (Outlook only) **Manifest type** — choose unified manifest unless a needed feature is
   XML-only.

Non-interactive / scripted scaffold:
```
yo office --projectType taskpane --name "my-addin" --host excel --ts true --skip-install
```
`--details` lists all flags. After `--skip-install`, run `npm install` in the project folder.

The generated project uses **webpack** + `webpack-dev-server` (HTTPS localhost, hot
reload), a TS→ES5 transpiler, and ships a working "Hello World" task pane.

## 6. Local development workflow

Run these from inside the add-in's project folder:

| Command | Purpose |
| --- | --- |
| `npm start` | Build, start the dev server, **and sideload** the add-in into the desktop Office app chosen in the manifest. |
| `npm run start:web` | Start for Office on the web (you supply the document URL). |
| `npm stop` | Stop the dev server and remove the sideloaded add-in. |
| `npm run build` | Production bundle into `dist/`. |
| `npm run dev-server` | Dev server only, no sideload. |
| `npm run lint` / `npm run lint:fix` | office-addin-lint (ESLint + Office rules). |
| `npm run validate` | Validate the manifest. |

Sideloading details and manual sideload steps (Windows registry share, Mac, web,
Teams/M365) are at `/office/dev/add-ins/testing/test-debug-office-add-ins`.

**First run of the dev server** installs a local CA certificate
(`office-addin-dev-certs`); accept the prompt or run `npx office-addin-dev-certs install`.

### Debugging

- **Task pane / commands**: `npm start` attaches the debugger; or use the browser
  devtools of the embedded webview (`npm run start:desktop -- --debug-method web`).
- **Script Lab** (free add-in from Microsoft Marketplace) — use it to prototype and
  verify Office.js snippets interactively inside Excel/Word/PowerPoint before porting
  code into the project.
- Runtime/logging: `office-addin-debugging` handles start/stop; check its output for
  sideload failures.

## 7. Office.js programming model

### Loading the library

Reference the CDN in the `<head>` of every add-in HTML page (never bundle it, never
self-host):

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
```

Preview APIs: use `.../lib/beta/hosted/office.js` (never ship `beta` to production).

### Initialization

All Office.js access must wait for the host to be ready:

```ts
Office.onReady((info) => {
  // info.host (Office.HostType), info.platform
  document.getElementById("run")!.onclick = run;
});
```

Do not call Office APIs at module top level.

### Two API models

- **Application-specific APIs** (`Excel`, `Word`, `PowerPoint`, `OneNote`) — strongly
  typed, promise-based, **batched**. Prefer these when the host has them.

  ```ts
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address,values");          // queue a read
    await context.sync();                  // execute the batch
    console.log(range.address);
    range.format.fill.color = "yellow";    // queue a write
    await context.sync();
  });
  ```

  Rules: `load()` only the properties you use; minimize `context.sync()` calls (each is a
  round trip — the killer of add-in perf on the web); don't hold proxy objects across
  `Excel.run` calls; handle `OfficeExtension.Error` (check `.code`).

- **Common APIs** (`Office.context.*`) — callback-based, one operation per request. Use
  for cross-host features: `Office.context.document` (get/set selected data),
  `Office.context.ui` (`displayDialogAsync`), `Office.context.roamingSettings` /
  document settings, and **all Outlook mail APIs** (`Office.context.mailbox.item`).

### Requirement sets

Gate any API that isn't universally available:

```ts
if (Office.context.requirements.isSetSupported("ExcelApi", "1.7")) { /* ... */ }
```

Also declare the minimum set in the manifest (`<Requirements>` / `extensions.requirements`)
so the add-in isn't offered on hosts that can't run it. Per-host/version/platform
support matrix: `/javascript/api/requirement-sets`.

## 8. Extending the Office UI

- **Add-in commands** — custom ribbon tab/group/buttons and context-menu items, defined
  in the manifest (`<VersionOverrides>` in XML). A command either opens a task pane
  (ShowTaskpane) or runs a JS function (ExecuteFunction) in the commands runtime.
- **Task panes** — the main surface; a web page docked beside the document.
- **Content add-ins** — embedded in the document body (Excel/PowerPoint), for
  dashboards/visualizations.
- **Dialogs** — `Office.context.ui.displayDialogAsync(url, options)`. The initial URL
  must be same-origin as the add-in; use `messageParent` / `messageChild` to communicate.
  Dialogs are the standard pattern for auth pop-ups.

An add-in is not required to have UI (e.g. a Copilot-agent-only add-in).

## 9. Runtimes

- **Shared runtime** — one persistent JS context across task pane + commands + custom
  functions; needed for cross-surface state, ribbon enable/disable, lifecycle events.
- **JavaScript-only runtime** — used by Excel custom functions without a shared runtime;
  calculation-optimized, restricted API surface, different programming model
  (`CustomFunctions.associate`).
- Configure in the manifest; Yo Office asks which for custom-function projects.

## 10. Coding conventions for this repo

- **TypeScript** for all new add-ins. Keep `strict` on.
- Keep the existing Yo Office toolchain (webpack, office-addin-* CLIs) unless there's a
  strong reason — pick the `manifest only` template if you need a different stack.
- Never commit `dist/`, `node_modules/`, or generated dev certs. Each project keeps its
  own `.gitignore` from the template.
- Keep secrets out of the client bundle. Client-side config that must vary by environment
  goes through webpack `DefinePlugin` / `.env`, not hardcoded.
- Match Microsoft's Fluent UI / Office design language for task-pane UI where practical.
- Run `npm run lint` and `npm run validate` before committing an add-in.

## 11. Security & hosting requirements

- **HTTPS everywhere** — web app, redirects, icons, external calls. Self-signed certs are
  fine for localhost dev only.
- Request the **least** manifest permission that works (`ReadDocument` <
  `ReadWriteDocument`; Outlook has its own permission tiers). Permission increases force
  admin re-consent on update.
- For user identity, prefer **SSO** (`OfficeRuntime.auth.getAccessToken` / Nested App
  Auth) over a custom OAuth dialog flow.
- Any iframe inside the add-in that calls Office.js must have its domain listed in the
  manifest or the call fails with permission-denied.

## 12. Deployment

| Method | When |
| --- | --- |
| **Sideloading** (`npm start`, or manual) | Dev/test only. |
| **Integrated Apps portal** (M365 admin center) | Distribute internal add-ins to users/groups in the org — no client config. Primary path for PDG internal tools. |
| **Centralized Deployment** (Exchange PowerShell / admin) | Org distribution in sovereign/gov clouds, or Outlook where Integrated Apps isn't available. |
| **Microsoft Marketplace (AppSource)** | Public distribution; must pass Commercial Marketplace certification (works on all platforms your APIs support, support URL in manifest, valid GUID). |
| **Network share / SharePoint catalog** | Legacy XML-manifest-only dev/on-prem options. Not for Outlook, not for production. |

Reference: `/office/dev/add-ins/publish/publish`. Bump the manifest version on every
deploy.

## 13. Useful references

- Develop overview: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview>
- Yo Office: <https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview> · repo <https://github.com/OfficeDev/generator-office>
- Manifest reference: `/office/dev/add-ins/develop/add-in-manifests`
- Requirement sets / platform availability: `/javascript/api/requirement-sets`
- Excel / Word / PowerPoint / Outlook API refs: `/javascript/api/{excel|word|powerpoint|outlook}`
- Script Lab: <https://aka.ms/scriptlab>
- Test & debug: `/office/dev/add-ins/testing/test-debug-office-add-ins`
