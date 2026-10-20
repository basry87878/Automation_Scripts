# download_sheets_browser

A Python script that automatically downloads Excel files shared through SharePoint / OneDrive links and places the latest copy in a OneDrive-synced folder, so it is uploaded back to SharePoint without any manual work.

It is built for files that are shared with a **specific email address**, where a plain download link does not work because Microsoft asks you to verify your email first. You verify once in a real browser window, the session is saved, and every scheduled run reuses it in the background.

## How it works

1. **One-time login.** `--login` opens a visible Edge (or Chrome) window using a dedicated profile folder. For each sheet you sign in or enter the verification code, and the script tests the download before moving on.
2. **Headless re-test.** After the login the script reopens the browser in headless mode, exactly like a scheduled run, to prove the sign-in survived. It prints a summary per sheet: `OK`, `LOST` or `NOT verified`.
3. **Scheduled runs.** Each run, for every sheet in `SHEETS`:
   - Converts the share link into a direct download URL (`Doc.aspx` / `doc2.aspx` links become `download.aspx?UniqueId=...`, other links get `download=1`).
   - Downloads the file using the saved browser session, and checks it is a real `.xlsx` (a zip file starting with `PK`). If that fails, it falls back to letting the browser trigger the download.
   - Retries up to 3 times with a growing pause between attempts.
   - Saves a timestamped snapshot plus a `_latest` copy locally.
   - Copies the file into the synced folder (`SYNC_DIR`) under the configured file name, replacing the old one.
   - Deletes old snapshots and old failure screenshots.
4. **Safe by design.**
   - Every sheet gets its own browser instance, so a crash or hang only affects that one sheet.
   - If a download fails, the previous good file in SharePoint is left untouched.
   - Files are written to a temporary name and then swapped in, so a half-written file is never published.
   - A lock file prevents two runs from using the browser profile at the same time. A lock older than 30 minutes is treated as stale.

## Requirements

- Python 3.8 or later
- Microsoft Edge or Google Chrome installed (the script uses the installed browser, no extra browser download)
- The `playwright` package
- Windows is assumed for the default paths. The script itself is plain Python.

```
pip install playwright
```

## Configuration

Everything is set at the top of `download_sheets_browser.py`.

| Setting | Meaning |
| ------- | ------- |
| `SHEETS` | One entry per file: `"short_name": ("full share link", "file name in SYNC_DIR")`. The short name is also used for the local folder and snapshot names. |
| `SYNC_DIR` | Local OneDrive-synced folder where the file is replaced. |
| `BROWSER_CHANNEL` | `"msedge"` or `"chrome"`. |
| `START_HOUR`, `END_HOUR` | Hours the script is allowed to run (inclusive). The default `8, 16` means 08:00 to 16:59. |
| `RETRIES`, `TIMEOUT_MS` | Download attempts per sheet and the timeout per attempt. |
| `SNAPSHOT_RETENTION_DAYS` | Timestamped snapshots and failure screenshots older than this are deleted (default 3). |
| `LOCK_MAX_AGE` | Seconds after which a lock file counts as stale (default 30 minutes). |

Example entry:

```python
SHEETS = {
    "My Report": (
        "https://<tenant>.sharepoint.com/:x:/r/personal/<user>/_layouts/15/Doc.aspx?sourcedoc=%7B<GUID>%7D&file=My%20Report.xlsx",
        "My Report.xlsx"),
}
```

Use the original link from the sharing email when you have it.

## Usage

First time, and whenever the session expires:

```
python download_sheets_browser.py --login
```

For each sheet, finish signing in in the browser window, then press Enter in the terminal to test the download (or type `s` to skip that sheet). When Microsoft asks "Stay signed in?", choose **Yes**.

Normal run (what the scheduler calls):

```
python download_sheets_browser.py
```

| Option | What it does |
| ------ | ------------ |
| `--login` | Interactive one-time verification in a visible browser. |
| `--force` | Ignore the `START_HOUR` / `END_HOUR` time window (useful for testing). |
| `--quick` | Testing mode: 1 attempt and a 30 second timeout per sheet, so failures show up fast. |
| `--only TEXT` | Only handle sheets whose name contains `TEXT`, for example `--login --only SAR` to redo a single sheet. |

Options can be combined, for example `python download_sheets_browser.py --force --quick`.

## Run it every hour with Task Scheduler

```
schtasks /create /tn "Download Sheets" /sc hourly /mo 1 /st 08:00 /f /tr "python C:\path\to\download_sheets_browser.py"
```

Notes:

- The script only works inside its allowed time window, so an all-day hourly schedule is fine. Outside the window it logs `nothing to do` and exits.
- If `python` is not on the PATH for the task, use the full path to `python.exe`.
- Run the task as the **same Windows user** who ran `--login`. Chromium-based browsers encrypt saved cookies per Windows user.
- The PC must be on and the user logged in.

## Output

Everything is created next to the script:

```
downloads/
  <name>/<name>_YYYYMMDD_HHMM.xlsx   hourly snapshots (deleted after 3 days)
  <name>/<name>_latest.xlsx          always the newest copy
  download.log                       full run log
  debug/fail_<name>_<time>.png       screenshot taken when a download fails
browser_profile/                     saved browser session
```

The exit code is `0` only when every sheet was downloaded and synced. Otherwise it is `1`, and `2` if `--only` matched no sheet.

## Troubleshooting

| Message or symptom | Meaning and fix |
| ------------------ | --------------- |
| `SESSION EXPIRED` | The saved sign-in is no longer valid. Run `--login` again, or `--login --only "<name>"` for one sheet. |
| `LOST after closing the browser` | The sign-in worked while the window was open but was not kept. Choose **Yes** on "Stay signed in?" and try again. |
| `another run is still in progress` | A previous run or login holds the lock. Wait for it to finish. If nothing is running, delete `downloads/.run.lock`. |
| `sync skipped: ... is open in another program` | The target file is open in Excel. The timestamped snapshot is still saved, and the next run will replace the file. |
| `sync skipped: folder not found` | `SYNC_DIR` is wrong or OneDrive is not running. |
| `outside 08:00-16:59 window` | The run was outside the allowed hours. Use `--force` to test. |
| A sheet keeps failing | Open the screenshot in `downloads/debug/` to see what the browser was showing, and check `download.log`. |

## Security notes

- The `browser_profile/` folder contains your saved Microsoft session. Treat it like a password.
- Share links contain access tokens and may contain email addresses. Do not publish real links.
- Do not commit `browser_profile/` or `downloads/`. Add this to `.gitignore`:

```
browser_profile/
downloads/
```

## Limitations

- Only `.xlsx` files are supported.
- It can only download files your account can already open through the shared link.
- If Microsoft changes the sign-in or sharing pages, `--login` may need to be repeated.
