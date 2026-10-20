# Madar_Tamer

A PowerShell script that automatically refreshes two report files from Outlook.

It finds the newest report email for each sender, saves the Excel attachment under a fixed file name, and replaces the old copy in a OneDrive-synced folder. OneDrive then uploads the new file to SharePoint, so everything that reads from that SharePoint location always sees the latest version.

## How it works

For each configured report, the script:

1. Connects to the classic Outlook desktop app and triggers a Send/Receive, then waits so the newest emails are downloaded.
2. Looks in two places: a mail folder (for example an archive folder) and the Inbox.
3. Picks the newest email whose subject matches the configured filter and that has a real file attachment (`.xlsx`, `.xlsm`, `.xlsb`, `.xls` or `.csv`).
4. Saves the attachment to a temporary file first, so a failed save never leaves a half-written file behind.
5. Compares the new file with the existing one. If they are identical, it skips the copy so OneDrive does not re-upload for nothing. Otherwise it replaces the existing file.
6. Writes everything it did to a log file next to the script.

Report folders and names are set in one list at the top of the script:

| Setting       | Meaning                                                  |
| ------------- | -------------------------------------------------------- |
| `Folder`      | Name of the Outlook folder to search (any mailbox/store) |
| `NewName`     | File name to save as (extension is kept from the email)  |
| `SubjectLike` | Subject filter, `*` matches any text                     |

## Requirements

- Windows with **classic Outlook** (desktop app), open and signed in. The "new Outlook" does not support the COM automation this script uses.
- Windows PowerShell 5.1 or later.
- The destination folder must be synced by the OneDrive client.

## Setup

1. Copy `Madar_Tamer.ps1` to a folder on your PC, for example `E:\Projects\`.
2. Open the script and edit the settings at the top:
   - `$Destination`: the OneDrive-synced folder where the files should be replaced.
   - `$Jobs`: one entry per report (`Folder`, `NewName`, `SubjectLike`).
   - Optional: `$AllowedExtensions`, `$ScanLimit`, `$SyncWaitSeconds`, `$StartupWaitSeconds`.
3. Run it once by hand to test (see below).

## Usage

Run manually and keep the window open to read the output:

```powershell
powershell -ExecutionPolicy Bypass -File E:\Projects\Madar_Tamer.ps1 -Pause
```

Expected output per report:

- `newest match ...` lines showing the email that was found in each location.
- `OK saved ...` when the file was replaced, or `SKIP already up to date` when nothing changed.

Running it twice in a row should give `SKIP` on the second run.

## Run it automatically with Task Scheduler

Example: every hour at minute 5, from 8:05 AM to 8:05 PM, every day:

```
schtasks /create /tn "Madar Tamer" /sc daily /st 08:05 /ri 60 /du 12:04 /f /tr "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File E:\Projects\Madar_Tamer.ps1"
```

The duration is set to 12 hours 4 minutes so the final 8:05 PM run is included.

On a laptop, allow the task to run on battery and to catch up on missed runs:

```powershell
$s = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Set-ScheduledTask -TaskName "Madar Tamer" -Settings $s
```

Test the scheduled task without waiting:

```
schtasks /run /tn "Madar Tamer"
```

Because it runs hidden, check `Madar_Tamer.log` for the result.

## Log

The script appends to `Madar_Tamer.log` in the same folder as the script. Check it first whenever something looks wrong. It shows which folders were searched, which email was picked, and any errors.

## Troubleshooting

| Problem | Likely cause and fix |
| ------- | -------------------- |
| Files are not updated, newest email in the log is old | Outlook had not downloaded the new mail. Keep Outlook running and connected, and make sure "Work Offline" is off. |
| `WARNING: Outlook is in Offline mode` in the log | Turn off Work Offline on the Send/Receive tab. |
| `Could not connect to Outlook` | Outlook is not installed, is the "new Outlook", or runs with different permissions than PowerShell (for example one is run as administrator and the other is not). |
| `no match in newest N emails` | The subject filter does not match, or the email has a link instead of a real attachment. The log shows the 3 newest subjects to help you adjust `SubjectLike`. |
| `FAILED` with a file in use message | The target file is open in Excel. Close it, and the next run will replace it. |
| Task does not run | The PC was asleep, or nobody was logged in. The task only runs for a logged-in user. |

## Notes and limitations

- Only real file attachments are supported. A link to a file in OneDrive or SharePoint is not downloaded.
- The script reads Outlook's local copy of the mailbox, so Outlook must be running and connected for new emails to be seen.
- The saved file keeps the email's received time as its modified date.
- If the existing file has a different extension than the new attachment, you will end up with two files.
- The script is read-only toward your mailbox. It never moves, deletes or edits emails.

How to test? Open powershell and run the following command:
powershell -ExecutionPolicy Bypass -File E:\Projects\Madar_Tamer.ps1 -Pause
