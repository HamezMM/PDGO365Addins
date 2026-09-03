# Installing the SheetToTxt Excel add-in

**What it does:** adds a **Home ▸ Export Sheet to .txt** button to Excel that saves the
current worksheet as a tab-delimited `.txt` file — next to the workbook, or via a
*Save As* prompt for SharePoint/cloud files.

**You need:** Windows, Excel (Microsoft 365 desktop). No admin rights. ~1 minute.

---

## Install

1. **Open the add-in folder.** In File Explorer go to:

   **Peake Design Group ▸ Peake Design - Documents ▸ SOFTWARE RESOURCES ▸ O365 ▸ SheetToTxt**

   If you don't see it, open it in the browser
   ([SharePoint link](https://chesapeakeud.sharepoint.com/sites/Peake/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FPeake%2FShared%20Documents%2FSOFTWARE%20RESOURCES%2FO365%2FSheetToTxt))
   and click **Sync** (or **Add shortcut to OneDrive**), then come back to Explorer.

2. **Run the installer.** Right-click **`Install-SheetToTxt.ps1`** ▸ **Run with PowerShell**.

   If that option is missing or nothing happens, open **PowerShell** from the Start menu and paste:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Peake Design Group\Peake Design - Documents\SOFTWARE RESOURCES\O365\SheetToTxt\Install-SheetToTxt.ps1"
   ```

3. **Approve it.** A **Microsoft Office Customization Installer** window appears — click
   **Install**. (This happens once.)

4. **Restart Excel.** Close every Excel window, reopen it. The **Export Sheet to .txt**
   button is on the **Home** tab, in a group called **Sheet to TXT**.

---

## Using it

1. Open a workbook and select the sheet you want to export.
2. **Home ▸ Export Sheet to .txt.**
   - Saved local / OneDrive-synced workbook → writes `<WorkbookName>.txt` beside it.
   - SharePoint / cloud-only / unsaved workbook → asks you where to save the `.txt`.
3. An existing `.txt` with the same name is overwritten.

The button is greyed out until a worksheet is active (open a workbook first).

---

## Updates

Automatic. When a new version is published, Excel picks it up the next time it starts —
nothing for you to do.

---

## Troubleshooting

| Problem | Fix |
| --- | --- |
| "running scripts is disabled on this system" | Use the full `powershell -ExecutionPolicy Bypass -File "..."` command in step 2. |
| No **Install** prompt / button never appears | Re-run the installer. Then in Excel: **File ▸ Options ▸ Add-ins**, set **Manage: COM Add-ins ▸ Go**, tick **SheetToTxt**. Also check **Manage: Disabled Items**. |
| Button shows but stays greyed out | Open or create a workbook — it enables when a worksheet is active. |
| "Publisher cannot be verified" every launch | The certificate isn't trusted. Re-run `Install-SheetToTxt.ps1` (it adds the trust). |
| Can't find the SharePoint folder | Ask in the team channel for access to **Peake Design ▸ SOFTWARE RESOURCES ▸ O365**. |

## Uninstall

Run **`Uninstall-SheetToTxt.ps1`** from the same folder, then restart Excel.

Questions: James / the team channel.
