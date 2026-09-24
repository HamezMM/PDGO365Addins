# Installing the BracketFields Word add-in

**What it does:** adds a **Home ▸ Fill [Fields]** button to Word. It finds every
`[bracketed]` placeholder in the document (body, tables, headers/footers, footnotes, text
boxes), walks you through each unique one so you can type its value, then replaces every
occurrence — brackets included — in one go.

**You need:** Windows, Word (Microsoft 365 desktop). No admin rights. ~1 minute.

---

## Install

1. **Open the add-in folder.** In File Explorer go to:

   **Peake Design Group ▸ Peake Design - Documents ▸ SOFTWARE RESOURCES ▸ O365 ▸ BracketFields**

   If you don't see it, open it in the browser
   ([SharePoint link](https://chesapeakeud.sharepoint.com/sites/Peake/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FPeake%2FShared%20Documents%2FSOFTWARE%20RESOURCES%2FO365%2FBracketFields))
   and click **Sync** (or **Add shortcut to OneDrive**), then come back to Explorer.

2. **Run the installer.** Right-click **`Install-BracketFields.ps1`** ▸ **Run with PowerShell**.

   If that option is missing or nothing happens, open **PowerShell** from the Start menu and paste:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Peake Design Group\Peake Design - Documents\SOFTWARE RESOURCES\O365\BracketFields\Install-BracketFields.ps1"
   ```

3. **Approve it.** A **Microsoft Office Customization Installer** window appears — click
   **Install**. (This happens once.)

4. **Restart Word.** Close every Word window, reopen it. The **Fill [Fields]** button is on
   the **Home** tab, in a group called **Bracket Fields**.

---

## Using it

1. Open a document containing placeholders like `[Client Name]`, `[Project Address]`, `[Date]`.
2. **Home ▸ Fill [Fields].** A window lists each unique field once, with how many times it
   appears and a snippet of surrounding text.
3. Type a value and press **Enter** to move to the next field (or click any field in the
   list). Leave a field blank to keep it as-is.
4. On the last field press **Enter** / **Finish** (or **Replace All** at any time).
5. Changed your mind? **Ctrl+Z** once undoes the whole fill.

Good to know:
- `[Date]` and `[date]` are treated as **different** fields — matching is exact.
- The replacement text keeps the formatting of the placeholder it replaces.
- Protected documents (Review ▸ Restrict Editing) must be unprotected first.

---

## Updates

Automatic. When a new version is published, OneDrive syncs it to your PC and Word picks
it up the next time it starts — nothing for you to do. Keep the folder synced.

---

## Troubleshooting

| Problem | Fix |
| --- | --- |
| "BracketFields.vsto not found" | OneDrive hasn't finished syncing the folder. Wait for the green ticks (or right-click the folder ▸ **Always keep on this device**), then re-run. |
| "running scripts is disabled on this system" | Use the full `powershell -ExecutionPolicy Bypass -File "..."` command in step 2. |
| No **Install** prompt / button never appears | Re-run the installer. Then in Word: **File ▸ Options ▸ Add-ins**, set **Manage: COM Add-ins ▸ Go**, tick **BracketFields**. Also check **Manage: Disabled Items**. |
| Button is greyed out | Open or create a document first. |
| "No [bracketed] fields were found" | The document has no `[...]` placeholders on a single line. Brackets split across lines or empty `[]` are ignored. |
| "Publisher cannot be verified" every launch | The certificate isn't trusted. Re-run `Install-BracketFields.ps1` (it adds the trust). |
| Can't find the SharePoint folder | Ask in the team channel for access to **Peake Design ▸ SOFTWARE RESOURCES ▸ O365**. |

## Uninstall

Run **`Uninstall-BracketFields.ps1`** from the same folder, then restart Word.

Questions: James / the team channel.
