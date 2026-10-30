# G Nexus - How to use

**Google Drive Migration Engine.** G Nexus backs up the Google Drive of several Google Workspace accounts to your PC: My Drive, Shared with me and Shared drives. It only copies. Nothing is ever changed or deleted on Google Drive.

## 1. First time on a new PC

1. Install rclone: open a command prompt and run `winget install Rclone.Rclone`, then close and reopen the window.
2. Put the G Nexus files in any folder. That folder only holds the app.
3. Double-click **G Nexus.bat**. On the first run it asks where the backups go: the **backup folder**.
4. Everything G Nexus makes goes into the backup folder: one folder per account, `_logs`, and `_gnexus` with `settings.json`, `accounts.csv` and `session.json` (made from the example files the first time). To use another backup folder later, press **Settings** or **Change...**. The settings and accounts are copied along.

## 2. Connect your own Google app

Do this once per organization. It takes about 10 minutes. Without it, rclone uses a shared Google app that is overloaded and gives "Quota exceeded" errors. Use a Workspace admin account.

1. [Create a project](https://console.cloud.google.com/projectcreate): give it a name and press **Create**.
2. [Turn on the Google Drive API](https://console.cloud.google.com/apis/library/drive.googleapis.com): press **Enable**.
3. [Branding](https://console.cloud.google.com/auth/branding): press **Get started**.
   - App name: `G Nexus`, support email: yours.
   - Audience: **Internal** when all accounts are in the same Workspace (no 7-day limit, no test users needed).
   - Or **External**: then add every account under *Audience > Test users*. Sign-ins expire after 7 days.
4. [Clients](https://console.cloud.google.com/auth/clients): **Create client**, Application type **Desktop app**, then **Create**.
5. In G Nexus press **Settings**, paste the **Client ID** and **Client secret** under *Google app*, and press **Save**.
6. Press **Check sign-in**. Every account shows `OLD SIGN-IN`. Press **Sign in all that need it**.

> If a sign-in says "blocked by your administrator": in [admin.google.com](https://admin.google.com) go to *Security > API controls > App access control* and trust the Client ID.

## 3. Accounts

- **Accounts > Add**: Name, Email, Password and Folder. The password is only a reminder shown during sign-in. rclone never uses it.
- **Edit list** opens `_gnexus\accounts.csv` (in the backup folder) to change or remove accounts.
  Columns: `Name, Email, Password, Folder` (`Remote` and `Legacy` are optional, leave them empty). An empty Folder is made from the Name.

## 4. Signing in

- Select an account and press **Sign in** (or **Sign in all that need it**).
- The output box shows the email and password, then the browser opens.
- In the browser choose **Use another account** if Google picks the wrong one.
- If Google says "Google hasn't verified this app", click **Continue**.
- When the browser says **Success!**, close the tab and go back to G Nexus.
- G Nexus checks the signed-in email. `WRONG ACCOUNT` means: sign in again.

## 5. Daily use

Click the account cards to select them (gold border), then press a button:

| Button | What it does |
|---|---|
| **Download** | Copies everything, then checks it byte by byte when *Verify after download* is ticked |
| **Sizes on Drive** | How much there is online (Download does not measure first, so it starts right away) |
| **Verify** | Byte check (size + MD5) of what is already downloaded |
| **SHA-256** | Checksum list to check a copy later |
| **Clean Mac files** | Deletes `._name`, `.DS_Store` and `__MACOSX` files already downloaded (no content, only Mac Finder info) |
| **Contacts + Calendar** | Opens Google's export pages for that account |
| **Download ALL** | Every signed-in account, one after another |

Options bar:

- **Parts**: My Drive, Shared with me, Shared drives, Trash. Only the ticked parts are downloaded, verified, measured and shown.
- **Shared with me**: only files from outside these accounts (recommended, after the cross-check) or all.
- **Verify after download**: byte check right after each download.
- **Overwrite files already downloaded**: off = only new or changed files.
- **Skip Mac junk files** (on by default): leaves out `._name`, `.DS_Store` and `__MACOSX` files that a Mac adds next to real files.
- **Skip shortcuts** (on by default): Drive shortcuts point to files or folders that live elsewhere, so their content is downloaded a second time. Tick it to leave them out. Keep it off if a shortcut points to something outside these accounts.
- **Same name twice on Drive**: Drive allows several files with the same name in one folder; Windows does not. Copies with the same content (same checksum) are saved once. When the contents differ, the first keeps the normal name and the others are saved as "name (dup 2)", "name (dup 3)", ...
- **Only files owned by the account** (on by default): My Drive downloads only the files this account owns. Files owned by other people are left out, because they belong to the owner's backup. It works in chunks: the files at the top of My Drive first, then one top folder at a time. For each chunk it reads the owners of that folder only and downloads it right away, so the download starts in seconds, even on a very big Drive. The output shows, per folder, how many files were left out. Careful with owners outside these accounts or already deleted: their files would be in no backup.
- **Table view**: the old table instead of the cards. **Show console window**: also shows the black window of the next action.

G Nexus remembers the ticked options and the selected accounts in `_gnexus\session.json` for next time. *Overwrite* always starts unticked.

Running:

- Everything runs inside G Nexus, with a progress bar and the output at the bottom.
- **Skip step** jumps over a step that is stuck. **Stop** cancels; files already downloaded are kept.
- Closing G Nexus while something runs asks whether to stop it. **Stop all (N)** shows how many backups run on this PC (also from closed windows) and stops them all. With G Nexus closed, double-click **Stop all backups.bat**.
- **Shared-with-me cross-check** finds shared files that belong to one of your own accounts (already backed up there) and writes `_REPORTE_Compartidos_conmigo.txt` in each account folder.

## 6. Before deleting any account

- [ ] The account shows **VERIFIED**.
- [ ] Size on Drive and Downloaded are close (Google Docs are converted, so sizes differ a bit).
- [ ] You read `_REPORTE_Compartidos_conmigo.txt` in its folder.
- [ ] The whole backup is copied to a second disk.

Google Forms, Sites, My Maps and Jamboard can't be downloaded by any tool.

## 7. Another organization

Use a new backup folder for each organization (**Settings > Save backups in**). Each backup folder has its own `_gnexus` with its own accounts and Google app. A new folder gets a copy of the current settings and accounts: change them there (step 2 and 3).

## 8. If something goes wrong

Press **Help > Report for support**. It saves `G-Nexus-report_<date>.zip` (and shows it in Explorer) with the logs, the errors, the saved results and system info. It has **no passwords and no Google tokens**. Send that file.

Errors are also written to `_logs\errors.log` (backup) and `_logs\app_errors.log` (window).

## Files

| File | What it is |
|---|---|
| `G Nexus.bat`, `Backup.bat` | Open G Nexus |
| `Backup-console.bat` | The same tool as a text menu |
| `_gnexus\settings.json` | In the backup folder. Saved by the Settings window: Google app, speed, speed limit, retries, document format, defaults |
| `_gnexus\accounts.csv` | In the backup folder. The accounts |
| `_gnexus\session.json` | In the backup folder. Ticked options and selected accounts |
| `_logs\` | In the backup folder. Logs, problem lists, saved results |
| `%APPDATA%\G Nexus\location.json` | Which backup folder G Nexus uses |
