# Installing the SpeakerRename Word add-in

**What it does:** adds a **Home ▸ Rename Speakers** button to Word. It finds every
`Speaker N` label in a transcript document (body, tables, headers/footers, footnotes,
comments, text boxes), walks you through each unique speaker number so you can type their
name, then replaces every occurrence of that speaker's label in one go.

**You need:** Windows, Word (Microsoft 365 desktop). No admin rights. ~1 minute.

---

## Install

1. **Open the add-in folder.** In File Explorer go to:

   **Peake Design Group ▸ Peake Design - Documents ▸ SOFTWARE RESOURCES ▸ O365 ▸ SpeakerRename**

   If you don't see it, open it in the browser
   ([SharePoint link](https://chesapeakeud.sharepoint.com/sites/Peake/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FPeake%2FShared%20Documents%2FSOFTWARE%20RESOURCES%2FO365%2FSpeakerRename))
   and click **Sync** (or **Add shortcut to OneDrive**), then come back to Explorer.

2. **Run the installer.** Right-click **`Install-SpeakerRename.ps1`** ▸ **Run with PowerShell**.

   If that option is missing or nothing happens, open **PowerShell** from the Start menu and paste:

   ```powershell
   powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Peake Design Group\Peake Design - Documents\SOFTWARE RESOURCES\O365\SpeakerRename\Install-SpeakerRename.ps1"
   ```

3. **Approve it.** A **Microsoft Office Customization Installer** window appears — click
   **Install**. (This happens once.)

4. **Restart Word.** Close every Word window, reopen it. The **Rename Speakers** button is
   on the **Home** tab, in a group called **Speaker Rename**.

---

## Using it

1. Open a transcript document containing labels like `Speaker 1`, `Speaker 2`, `Speaker 3`.
2. **Home ▸ Rename Speakers.** A window lists each unique speaker number once, with how
   many times it appears and a snippet of surrounding text.
3. Type a name and press **Enter** to move to the next speaker (or click any speaker in the
   list). Leave a speaker blank to keep their label as `Speaker N`.
4. On the last speaker press **Enter** / **Finish** (or **Rename All** at any time).
5. Changed your mind? **Ctrl+Z** once undoes the whole rename.

Good to know:
- Speakers are grouped by **number**, not by exact spacing: `Speaker  3` (extra space) and
  `Speaker 3` count as the same speaker and are both replaced.
- The replacement text keeps the formatting of the label it replaces.
- Protected documents (Review ▸ Restrict Editing) must be unprotected first.

---

## Updates

Automatic. When a new version is published, OneDrive syncs it to your PC and Word picks
it up the next time it starts — nothing for you to do. Keep the folder synced.

---

## Troubleshooting

| Problem | Fix |
| --- | --- |
| "SpeakerRename.vsto not found" | OneDrive hasn't finished syncing the folder. Wait for the green ticks (or right-click the folder ▸ **Always keep on this device**), then re-run. |
| "running scripts is disabled on this system" | Use the full `powershell -ExecutionPolicy Bypass -File "..."` command in step 2. |
| No **Install** prompt / button never appears | Re-run the installer. Then in Word: **File ▸ Options ▸ Add-ins**, set **Manage: COM Add-ins ▸ Go**, tick **SpeakerRename**. Also check **Manage: Disabled Items**. |
| Button is greyed out | Open or create a document first. |
| "No 'Speaker N' labels were found" | The document has no `Speaker` + number text. Check the exact spelling/capitalization used in the transcript. |
| "Publisher cannot be verified" every launch | The certificate isn't trusted. Re-run `Install-SpeakerRename.ps1` (it adds the trust). |
| Can't find the SharePoint folder | Ask in the team channel for access to **Peake Design ▸ SOFTWARE RESOURCES ▸ O365**. |

## Uninstall

Run **`Uninstall-SpeakerRename.ps1`** from the same folder, then restart Word.

Questions: James / the team channel.
