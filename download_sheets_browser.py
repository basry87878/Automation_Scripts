#!/usr/bin/env python3
"""
Hourly downloader for SharePoint/OneDrive files shared to a *specific email*.

You verify your email ONCE in a real browser window; the session is saved in a
dedicated profile folder, and the scheduled runs reuse it (headless).

Setup (once):
    pip install playwright            # uses your installed Edge, no extra download
    python download_sheets_browser.py --login

Scheduled run (every hour; only works between START_HOUR and END_HOUR):
    python download_sheets_browser.py
    python download_sheets_browser.py --force     # ignore the time window (testing)
    python download_sheets_browser.py --force --quick   # testing: 1 attempt, 30 s timeout
    python download_sheets_browser.py --login --only SAR   # redo just one sheet (name contains "SAR")

When the saved session expires the log says "SESSION EXPIRED" -> run --login again.
Only one run (or --login) can use the browser profile at a time; a second one
logs "another run is still in progress" and exits.

Output:
    downloads/<name>/<name>_YYYYMMDD_HHMM.xlsx   (hourly snapshots)
    downloads/<name>/<name>_latest.xlsx          (always the newest copy)
    downloads/download.log
    downloads/debug/fail_<name>_<time>.png       (screenshot when a download fails)
"""
import argparse
import logging
import os
import sys
import time
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from playwright.sync_api import sync_playwright

# ------------------------------------------------------------------ config --
# short_name -> FULL share link. Add the future sheets here, one line each.
# short_name -> (FULL share link, sync filename inside SYNC_DIR)
SHEETS = {
    "APG Control Tower Report": (
        "https://abudawood-my.sharepoint.com/:x:/r/personal/mohamedidrisai_abudawood_com/_layouts/15/Doc.aspx?sourcedoc=%7B96BE6214-B316-440E-9563-04A316A86555%7D&file=APG%20Control%20Tower%20Report.xlsx&fromShare=true&action=default&mobileredirect=true",
        "APG Control Tower Report.xlsx"),
    "SAR X P&G Control Tower": (
        # Original share link (what the email contains). Same share token as the old
        # guestaccess.aspx link below, which is just where this one redirects to.
        # old: https://bahri1-my.sharepoint.com/:x:/r/personal/c-kabdul_bahri_sa/_layouts/15/guestaccess.aspx?email=basry.ab%40pg.com&e=PVoOe1&share=IQBCOOuVRxFPSYzEiQyFOnjmAStcc7ai0vw90ekAehpWSx0
        "https://bahri1-my.sharepoint.com/:x:/g/personal/c-kabdul_bahri_sa/IQBCOOuVRxFPSYzEiQyFOnjmAStcc7ai0vw90ekAehpWSx0?email=basry.ab%40pg.com&e=PVoOe1",
        "SAR X P&G Control Tower.xlsx"),
    "pg_ksa_control_tower": (
        # Original share form, built the same way as the SAR link (token moves from
        # share=... into the path). If you still have the email, its link is better.
        # old: https://ahktrans1-my.sharepoint.com/:x:/r/personal/altaf_noor_ahk-trans_com/_layouts/15/guestaccess.aspx?e=4%3AKolVrB&at=9&share=EQguzDx_0-5ct4TrVEe19noBm3fRmuEp1V14wifbSkALew
        "https://ahktrans1-my.sharepoint.com/:x:/g/personal/altaf_noor_ahk-trans_com/EQguzDx_0-5ct4TrVEe19noBm3fRmuEp1V14wifbSkALew?e=4%3AKolVrB&at=9",
        "P&G KSA Control Tower Rep.xlsx"),
    "CONTROL SHEET P&G _25": (
        "https://my.maerskgroup.com/:x:/r/personal/karim_kamal_maersk_com/_layouts/15/doc2.aspx?sourcedoc=%7B53116E0C-4079-4F9F-9FD4-64FFEFEDDB5E%7D&file=CONTROL%20SHEET%20P%26G%20_25.xlsx&action=default&mobileredirect=true",
        "CONTROL SHEET P&G _25.xlsx"),
    "Nahdi Control Tower": (
        "https://ntcosa.sharepoint.com/:x:/r/sites/ITZone/_layouts/15/Doc.aspx?sourcedoc=%7Bba426304-ec09-424e-a8d3-b40ebc4ae7d4%7D&action=default&CID=21c781fe-55c5-4c32-b510-b263427845e9&_SRM=2%3AE%3A33",
        "Nahdi Control Tower.xlsx"),
}

BASE_DIR = Path(__file__).resolve().parent
OUT_DIR = BASE_DIR / "downloads"
PROFILE_DIR = BASE_DIR / "browser_profile"  # saved login session lives here
BROWSER_CHANNEL = "msedge"                  # or "chrome"; both use the installed browser
START_HOUR, END_HOUR = 8, 16                # inclusive -> runs 08:00 ... 16:59
RETRIES = 3
TIMEOUT_MS = 60_000

LOCK_FILE = OUT_DIR / ".run.lock"           # prevents two runs sharing the browser profile
LOCK_MAX_AGE = 30 * 60                      # seconds; an older lock is treated as stale
DEBUG_DIR = OUT_DIR / "debug"               # failure screenshots (auto-deleted like snapshots)

# -- SharePoint sync copy ------------------------------------------------------
# Local path of the synced SharePoint folder (OneDrive syncs it up automatically).
# The downloaded file is copied here with SYNC_FILENAME, REPLACING the old one.
SYNC_DIR = Path(r"C:\Users\Test\OneDrive - Procter and Gamble\Saudi PD - Transportation\Last Mile Operations Officer\Source\Transporters Reports")
SNAPSHOT_RETENTION_DAYS = 3   # timestamped backups older than this are deleted
# -----------------------------------------------------------------------------


class SessionExpired(Exception):
    pass


# -------------------------------------------------------------------- lock --
def acquire_lock() -> bool:
    """Atomically create the lock file. False if another run holds a fresh lock."""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for _ in range(2):
        try:
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                age = time.time() - LOCK_FILE.stat().st_mtime
            except FileNotFoundError:
                continue                     # released in between -> try again
            if age < LOCK_MAX_AGE:
                return False
            try:
                LOCK_FILE.unlink()           # stale lock from a crashed run
            except OSError:
                return False
            continue
        with os.fdopen(fd, "w") as fh:
            fh.write(str(os.getpid()))
        return True
    return False


def refresh_lock() -> None:
    """Keep a long run's lock from looking stale."""
    try:
        os.utime(LOCK_FILE)
    except OSError:
        pass


def release_lock() -> None:
    try:
        LOCK_FILE.unlink()
    except OSError:
        pass


def to_direct_download(link: str):
    """Doc.aspx / doc2.aspx links -> .../_layouts/15/download.aspx?UniqueId=<file id>"""
    parts = urlsplit(link)
    m = re.search(r"sourcedoc=%7B([0-9a-fA-F-]+)%7D", link, re.I)
    if not m or not re.search(r"/_layouts/15/doc2?\.aspx", parts.path, re.I):
        return None
    base = parts.path.split("/_layouts/")[0]      # /:x:/r/personal/<user>
    base = re.sub(r"^/:[a-z]:/[a-z]/", "/", base)  # -> /personal/<user>
    return urlunsplit((parts.scheme, parts.netloc,
                       f"{base}/_layouts/15/download.aspx",
                       f"UniqueId={m.group(1)}", ""))

def with_download_flag(link: str) -> str:
    direct = to_direct_download(link)
    if direct:
        return direct
    parts = urlsplit(link)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["download"] = "1"
    return urlunsplit(parts._replace(query=urlencode(query)))


def launch(p, headless: bool):
    return p.chromium.launch_persistent_context(
        str(PROFILE_DIR), channel=BROWSER_CHANNEL, headless=headless, accept_downloads=True
    )


# ------------------------------------------------------------------- login --
def login() -> int:
    """Open a visible browser so you can do the email/code verification once."""
    if not acquire_lock():
        print(f"A download run is in progress (lock file: {LOCK_FILE}).\n"
              "Wait for it to finish and try again. If you're sure nothing is running, "
              "delete the lock file.")
        return 1
    passed, lost = [], []
    try:
        with sync_playwright() as p:
            # Phase 1: visible browser. Each sheet's download is TESTED before moving on.
            ctx = launch(p, headless=False)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            for name, (link, _sync_name) in SHEETS.items():
                if link.startswith("PASTE_"):
                    print(f"{name}: set the real share link in SHEETS first")
                    continue
                print(f"\n[{name}]")
                try:
                    page.goto(link)
                except Exception as exc:  # noqa: BLE001
                    print(f"   page did not finish loading ({str(exc).splitlines()[0]}); "
                          "carry on in the browser window")
                prompt = ("   Sign in / verify until the sheet is fully open, "
                          "then press Enter here (or type s to skip)... ")
                while True:
                    if input(prompt).strip().lower() == "s":
                        break
                    refresh_lock()
                    if check_download(ctx, name, link):
                        passed.append(name)
                        break
                    prompt = ("   Finish signing in in the browser window, then press Enter "
                              "to test again (or type s to skip)... ")
            ctx.close()

            # Phase 2: reopen HEADLESS, exactly like a scheduled run, to prove the
            # sign-ins survive closing the browser.
            if passed:
                print("\nBrowser closed. Re-testing the way the scheduled task will run...")
                time.sleep(2)
                ctx = launch(p, headless=True)
                try:
                    for name in passed:
                        if not check_download(ctx, name, SHEETS[name][0]):
                            lost.append(name)
                finally:
                    ctx.close()
    finally:
        release_lock()

    print("\n=== Summary ===")
    for name in SHEETS:
        if name in lost:
            status = "LOST after closing the browser (sign-in was not kept)"
        elif name in passed:
            status = "OK"
        else:
            status = "NOT verified"
        print(f"  {name}: {status}")
    if lost:
        print("\nThe LOST sheets worked while the window was open but not after it closed.\n"
              "When Microsoft asks 'Stay signed in?', choose Yes. If it still happens, "
              "tell me and I'll make the script save those cookies explicitly.")
    ready = len(passed) - len(lost) == len(SHEETS)
    if ready:
        print("\nAll sheets verified. Run:  python download_sheets_browser.py --force")
    return 0 if ready else 1


# ---------------------------------------------------------------- download --
def fetch_via_request(ctx, url: str) -> bytes:
    resp = ctx.request.get(url, timeout=TIMEOUT_MS)  # shares the browser's cookies
    return resp.body() if resp.ok else b""


LOGIN_HOSTS = ("login.microsoftonline.com", "login.live.com")
LOGIN_TITLE_HINTS = ("sign in", "verify", "log in", "sharing link validation")


def relogin_cmd(name: str) -> str:
    only = f' --only "{name}"' if name else ""
    return f"python download_sheets_browser.py --login{only}"


def looks_like_login(title: str, url: str) -> bool:
    return (any(h in url for h in LOGIN_HOSTS)
            or any(h in title.lower() for h in LOGIN_TITLE_HINTS))


def fetch_via_browser(ctx, url: str, name: str = "") -> bytes:
    """Fallback: let the browser trigger the download like a click would."""
    page = ctx.new_page()
    nav_err = None
    try:
        with page.expect_download(timeout=TIMEOUT_MS) as info:
            try:
                page.goto(url, timeout=TIMEOUT_MS)
            except Exception as exc:  # noqa: BLE001  ("Download is starting" is expected)
                nav_err = exc
        return Path(info.value.path()).read_bytes()
    except Exception:
        # No download started: record what the browser is actually showing.
        title, where, shot = "?", "?", None
        try:
            title = page.title()
            where = page.url.split("?")[0][:150]       # no query string: it holds share tokens
            DEBUG_DIR.mkdir(parents=True, exist_ok=True)
            safe = re.sub(r"[^\w.-]+", "_", name) or "sheet"
            shot = DEBUG_DIR / f"fail_{safe}_{datetime.now():%Y%m%d_%H%M%S}.png"
            page.screenshot(path=str(shot))
        except Exception:  # noqa: BLE001
            pass
        logging.warning("%s: no download started; page title=%r url=%s screenshot=%s%s",
                        name or "?", title, where, shot,
                        f" navigation_error={str(nav_err).splitlines()[0]}" if nav_err else "")
        if looks_like_login(title, where):
            raise SessionExpired(
                f"SESSION EXPIRED - the browser got a sign-in/verification page ({title!r}). "
                f"Run: {relogin_cmd(name)}"
            )
        raise
    finally:
        page.close()


def fetch_one(ctx, link: str, name: str = "") -> bytes:
    url = with_download_flag(link)
    last_err = None
    for attempt in range(1, RETRIES + 1):
        try:
            body = fetch_via_request(ctx, url)
            if not body.startswith(b"PK"):  # .xlsx is a zip -> starts with "PK"
                body = fetch_via_browser(ctx, url, name)
            if not body.startswith(b"PK"):
                raise SessionExpired(
                    "SESSION EXPIRED or download blocked - run: "
                    + relogin_cmd(name)
                )
            return body
        except SessionExpired:
            raise
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            logging.warning("attempt %d/%d failed: %s", attempt, RETRIES, exc)
            if attempt < RETRIES:
                time.sleep(5 * attempt)
    raise RuntimeError(f"all {RETRIES} attempts failed: {last_err}")


def check_download(ctx, name: str, link: str) -> bool:
    """Quick one-attempt download test, used by --login. Prints the result."""
    global RETRIES, TIMEOUT_MS
    saved = RETRIES, TIMEOUT_MS
    RETRIES, TIMEOUT_MS = 1, 30_000
    try:
        data = fetch_one(ctx, link, name)
        print(f"   OK - download works ({len(data) / 1024:.0f} KB)")
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"   NOT downloadable yet: {str(exc).splitlines()[0]}")
        return False
    finally:
        RETRIES, TIMEOUT_MS = saved


def save(name: str, data: bytes, stamp: datetime) -> Path:
    folder = OUT_DIR / name
    folder.mkdir(parents=True, exist_ok=True)

    dated = folder / f"{name}_{stamp:%Y%m%d_%H%M}.xlsx"
    dated.write_bytes(data)

    latest = folder / f"{name}_latest.xlsx"
    tmp = folder / f"{name}_latest.tmp"
    tmp.write_bytes(data)
    try:
        os.replace(tmp, latest)
    except PermissionError:
        logging.warning("%s is open in another program; kept timestamped copy only", latest.name)
        tmp.unlink(missing_ok=True)
    return dated


def sync_copy(name: str, sync_name: str, data: bytes) -> bool:
    """Copy the fresh download into the OneDrive-synced SharePoint folder.
    Returns True only if the file was really replaced."""
    if not SYNC_DIR or not SYNC_DIR.exists():
        logging.warning("%s: sync skipped: folder not found: %s", name, SYNC_DIR)
        return False
    target = SYNC_DIR / sync_name
    tmp = SYNC_DIR / (sync_name + ".tmp")
    try:
        tmp.write_bytes(data)
        os.replace(tmp, target)
        logging.info("%s: synced copy -> %s", name, target)
        return True
    except PermissionError:
        logging.warning("%s: sync skipped: %s is open in another program", name, target.name)
    except OSError as exc:
        logging.warning("%s: sync failed: %s", name, exc)
    tmp.unlink(missing_ok=True)
    return False


def cleanup_snapshots(name: str, now: datetime) -> None:
    """Delete timestamped snapshots older than SNAPSHOT_RETENTION_DAYS."""
    folder = OUT_DIR / name
    if not folder.exists():
        return
    cutoff = now.timestamp() - SNAPSHOT_RETENTION_DAYS * 86400
    removed = 0
    for f in folder.glob(f"{name}_2*.xlsx"):  # e.g. pg_ksa_control_tower_20261014_1256.xlsx
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
                removed += 1
        except OSError as exc:
            logging.warning("cleanup: could not delete %s: %s", f.name, exc)
    if removed:
        logging.info("%s: cleaned up %d snapshot(s) older than %d days",
                     name, removed, SNAPSHOT_RETENTION_DAYS)


def cleanup_debug(now: datetime) -> None:
    """Delete failure screenshots older than SNAPSHOT_RETENTION_DAYS."""
    if not DEBUG_DIR.exists():
        return
    cutoff = now.timestamp() - SNAPSHOT_RETENTION_DAYS * 86400
    for f in DEBUG_DIR.glob("fail_*.png"):
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
        except OSError:
            pass


def launch_with_retry(p, tries: int = 3):
    """Start the headless browser; a busy profile or a hiccup gets 3 chances."""
    last = None
    for i in range(1, tries + 1):
        try:
            return launch(p, headless=True)
        except Exception as exc:  # noqa: BLE001
            last = exc
            logging.warning("browser start attempt %d/%d failed: %s",
                            i, tries, str(exc).splitlines()[0])
            if i < tries:
                time.sleep(10)
    raise RuntimeError(f"could not start the browser: {str(last).splitlines()[0]}")


def close_quietly(ctx) -> None:
    if ctx is None:
        return
    try:
        ctx.close()
    except Exception:  # noqa: BLE001
        pass


def guarded(step: str, name: str, fn, *args):
    """Run one step; if it blows up, log it and carry on instead of aborting the sheet."""
    try:
        return fn(*args)
    except Exception as exc:  # noqa: BLE001
        logging.warning("%s: %s failed (continuing): %s", name, step, exc)
        return None


def run() -> int:
    now = datetime.now()
    downloaded = synced = failures = 0
    cleanup_debug(now)
    with sync_playwright() as p:
        for name, (link, sync_name) in SHEETS.items():
            refresh_lock()
            if link.startswith("PASTE_"):
                logging.error("%s: set the real share link in SHEETS first", name)
                failures += 1
                continue

            # 1) Download. Every sheet gets its OWN browser, so a crash, a hang or a
            #    start-up problem can only ever cost that one sheet.
            data, ctx = None, None
            try:
                ctx = launch_with_retry(p)
                data = fetch_one(ctx, link, name)
            except Exception as exc:  # noqa: BLE001
                logging.error("%s: FAILED - %s", name, exc)
                failures += 1
            finally:
                close_quietly(ctx)
            if data is None:
                continue          # the previous good file stays in SharePoint untouched
            downloaded += 1

            # 2) The three steps below are independent: none can stop the others.
            path = guarded("local backup", name, save, name, data, now)
            if path:
                logging.info("%s: saved %s (%.1f KB)", name, path.name, len(data) / 1024)
            if guarded("sync", name, sync_copy, name, sync_name, data):
                synced += 1
            guarded("cleanup", name, cleanup_snapshots, name, now)

    logging.info("run finished: %d downloaded, %d synced, %d failed",
                 downloaded, synced, failures)
    return 0 if (failures == 0 and synced == downloaded) else 1


def main() -> int:
    global RETRIES, TIMEOUT_MS, SHEETS
    parser = argparse.ArgumentParser()
    parser.add_argument("--login", action="store_true", help="one-time interactive verification")
    parser.add_argument("--force", action="store_true", help="ignore the time window")
    parser.add_argument("--quick", action="store_true",
                        help="testing: 1 attempt and 30 s timeout per sheet, so failures show up fast")
    parser.add_argument("--only", metavar="TEXT",
                        help="only handle sheets whose name contains TEXT, e.g. --login --only SAR")
    args = parser.parse_args()

    if args.quick:
        RETRIES, TIMEOUT_MS = 1, 30_000
    if args.only:
        SHEETS = {k: v for k, v in SHEETS.items() if args.only.lower() in k.lower()}
        if not SHEETS:
            print(f"No sheet name contains {args.only!r}.")
            return 2

    if args.login:
        return login()

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(OUT_DIR / "download.log", encoding="utf-8"),
                  logging.StreamHandler(sys.stdout)],
    )

    hour = datetime.now().hour
    if not args.force and not (START_HOUR <= hour <= END_HOUR):
        logging.info("outside %02d:00-%02d:59 window, nothing to do", START_HOUR, END_HOUR)
        return 0

    if not acquire_lock():
        logging.info("another run is still in progress, skipping (lock: %s)", LOCK_FILE)
        return 0
    try:
        return run()
    except Exception:  # noqa: BLE001
        logging.exception("run crashed unexpectedly")
        return 1
    finally:
        release_lock()


if __name__ == "__main__":
    sys.exit(main())