"""
fetch_osm_report.py — automated pull of the daily OSM "All Open Tasks by
Team" report attachment from Outlook, for the HRIS dashboard's morning
auto-refresh.

WHAT THIS DOES (and does NOT do):
  - Connects to Outlook via COM, looks in the specific mail-rule-filtered
    folder Inbox > Reports > OSM for today's report email from
    reports-prd-ldz@saasiteu.com, subject "Report: All Open Tasks by Team".
  - Saves its .xls/.xlsx attachment into the Downloads folder, using the
    SAME filename pattern import_osm_report.py already looks for
    ("All Open Tasks by Team*.xls" / "*.xlsx"), so the existing, UNCHANGED
    manual pipeline (import_osm_report.py / Update HRIS Dashboard.bat)
    picks it up exactly as if Kevin had downloaded it by hand.
  - Does NOT parse the spreadsheet, does NOT push to GitHub, does NOT touch
    data/tickets.json. That work stays entirely inside import_osm_report.py,
    unmodified. This script's only job is "get today's attachment into the
    folder the existing pipeline already watches."
  - Saves to a fixed, dedicated filename ("All Open Tasks by Team - auto.xls")
    distinct from anything Kevin downloads by hand, so an auto-fetch never
    silently overwrites a manual download with the same name, while still
    matching the existing glob pattern (find_report() in
    import_osm_report.py picks whichever matching file has the newest
    mtime, regardless of exact name) — so on any given day, whichever of
    "manual download" or "this morning's auto-fetch" is more recent wins,
    with zero change needed to the existing selection logic.

FAILS LOUDLY, ON PURPOSE, if:
  - Outlook COM cannot be reached after retries.
  - The Inbox > Reports > OSM folder does not exist (folder renamed/moved).
  - No email from the expected sender+subject was received TODAY. This is
    deliberate: silently reusing yesterday's leftover attachment would make
    the dashboard look freshly updated when it is not. A missing report is
    surfaced as a failure, not swallowed.
  - The matching email has no attachment, or its attachment's filename
    doesn't look like an "All Open Tasks by Team" export. This is the
    detection mechanism for "OSM changed the report's sender/subject/
    attachment format" — the exact silent-break scenario this script exists
    to avoid.

Exit code 0 = attachment saved successfully, safe to run the dashboard
update next. Any non-zero exit code = do NOT run the dashboard update this
morning; the caller (Run_HRIS_Auto_Refresh.bat) checks this and skips that
step, so a bad/missing report never gets silently "updated" over.

Usage:
    python fetch_osm_report.py             # real run — saves to Downloads
    python fetch_osm_report.py --dry-run PATH
                                            # verification run — saves to
                                            # PATH instead of the real
                                            # Downloads folder, so it can be
                                            # proven against the real
                                            # mailbox without touching the
                                            # live pipeline's input file.

Requires: pip install pywin32
"""

import os
import sys
import time
import argparse
from datetime import datetime

import win32com.client
import pywintypes

# ── Config ───────────────────────────────────────────────────────────────

OSM_SENDER_EMAIL = "reports-prd-ldz@saasiteu.com"
OSM_SUBJECT       = "Report: All Open Tasks by Team"
OSM_FOLDER_PATH   = ["Reports", "OSM"]   # under the default store's Inbox

DOWNLOADS_DIR       = r"C:\Users\admin\Downloads"
SAVED_FILENAME_BASE = "All Open Tasks by Team - auto"   # extension added from the real attachment

# Every run must print a clear timestamp so pasted console/log output is
# self-dating (standing rule, Kevin 12 Aug 2026 — a pasted traceback with no
# timestamp made an already-fixed incident look like a fresh one).
def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


# Late-binding COM connection with a bounded retry, matching the pattern
# already proven live in work-inbox/fetch_inbox.py (confirmed there twice,
# 2026-08-11: "Call was rejected by callee" is Outlook being transiently
# busy mid-sync / a modal dialog open, not a real fault — a retry a few
# minutes later succeeded cleanly both times). Reused here rather than
# re-derived, per begb0037admin/drew's own memory discipline.
def connect_to_outlook(max_attempts=3, retry_wait_seconds=45):
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            outlook_app = win32com.client.dynamic.Dispatch("Outlook.Application")
            mapi_ns = outlook_app.GetNamespace("MAPI")
            inbox = mapi_ns.GetDefaultFolder(6)  # olFolderInbox, touched now to
            # surface a busy-callee failure here rather than later.
            if attempt > 1:
                log(f"Outlook COM connection succeeded on attempt {attempt}/{max_attempts}.")
            return outlook_app, mapi_ns, inbox
        except pywintypes.com_error as e:
            last_error = e
            log(f"Outlook COM connection attempt {attempt}/{max_attempts} failed: {e}")
            if attempt < max_attempts:
                log(f"Outlook automation layer appears busy (transient). Waiting {retry_wait_seconds}s before retrying...")
                time.sleep(retry_wait_seconds)
    log(f"Outlook COM connection failed after {max_attempts} attempts. Giving up.")
    raise last_error


def find_osm_folder(inbox):
    folder = inbox
    path_so_far = "Inbox"
    for name in OSM_FOLDER_PATH:
        try:
            folder = folder.Folders(name)
        except pywintypes.com_error:
            sys.exit(
                f"ERROR: Outlook folder '{path_so_far}/{name}' not found. "
                f"The OSM report folder may have been renamed or moved — "
                f"check the mail rule and folder structure in Outlook."
            )
        path_so_far = f"{path_so_far}/{name}"
    log(f"Found folder: {path_so_far} ({folder.Items.Count} items)")
    return folder


def find_todays_report(folder):
    """Returns the MailItem for today's OSM report, or None if not found.

    Also returns diagnostic info about the most recent matching-sender
    email (regardless of date) so a failure message can say *why* — e.g.
    "the sender matched but it's from yesterday" vs "no email from this
    sender at all" are very different failure modes worth distinguishing.
    """
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # descending — newest first

    today = datetime.now().date()
    most_recent_from_sender = None

    checked = 0
    for msg in items:
        checked += 1
        if checked > 30:
            # The folder holds ~50 historical reports; today's (if present)
            # will always be at/near the top given descending sort. 30 is
            # generous headroom, not a tight fit.
            break
        try:
            if getattr(msg, "Class", None) != 43:  # olMail only
                continue
            sender = (getattr(msg, "SenderEmailAddress", "") or "").strip().lower()
            subject = (getattr(msg, "Subject", "") or "").strip()
            if sender != OSM_SENDER_EMAIL.lower():
                continue
            if most_recent_from_sender is None:
                most_recent_from_sender = msg
            if subject != OSM_SUBJECT:
                log(f"  (skipping — sender matched but subject was {subject!r}, expected {OSM_SUBJECT!r})")
                continue
            received = msg.ReceivedTime
            received_date = datetime(received.year, received.month, received.day).date()
            if received_date == today:
                return msg, most_recent_from_sender
            else:
                log(f"  (most recent matching email is from {received_date}, not today {today})")
                return None, most_recent_from_sender
        except Exception as e:
            log(f"  (error reading an item while scanning: {e})")
            continue

    return None, most_recent_from_sender


def save_attachment(msg, target_dir, dry_run_label=""):
    n = msg.Attachments.Count
    if n == 0:
        sys.exit(
            "ERROR: Today's OSM report email has NO attachment. "
            "This is exactly the kind of silent format change this script "
            "watches for — check the email in Outlook (Inbox/Reports/OSM) directly."
        )

    # Expect exactly one attachment named like the known report. If OSM
    # ever ships multiple attachments or renames it away from "All Open
    # Tasks by Team", fail loudly rather than guess.
    candidate = None
    for i in range(1, n + 1):
        att = msg.Attachments.Item(i)
        fname = att.FileName or ""
        ext = os.path.splitext(fname)[1].lower()
        if fname.lower().startswith("all open tasks by team") and ext in (".xls", ".xlsx"):
            candidate = att
            break

    if candidate is None:
        # Fall back to the first attachment but flag it loudly — this is
        # the detection path for "OSM renamed the file".
        att = msg.Attachments.Item(1)
        fname = att.FileName or "(unknown)"
        log(f"WARNING: no attachment matched the expected 'All Open Tasks by Team*.xls(x)' "
            f"naming pattern. Found {n} attachment(s), first is {fname!r}.")
        ext = os.path.splitext(fname)[1].lower()
        if ext not in (".xls", ".xlsx"):
            sys.exit(
                f"ERROR: attachment {fname!r} does not look like an Excel export "
                f"(extension {ext!r}) — refusing to save it as today's report. "
                f"Check the email directly — the report format may have changed."
            )
        candidate = att
    else:
        fname = candidate.FileName
        ext = os.path.splitext(fname)[1].lower()

    target_path = os.path.join(target_dir, f"{SAVED_FILENAME_BASE}{ext}")
    os.makedirs(target_dir, exist_ok=True)
    candidate.SaveAsFile(target_path)
    size = os.path.getsize(target_path)
    log(f"Saved attachment {fname!r} ({size} bytes) -> {target_path}{dry_run_label}")
    return target_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--dry-run", metavar="DIR", default=None,
        help="Save the attachment into DIR instead of the real Downloads "
             "folder, for verification without touching the live pipeline's input.",
    )
    args = parser.parse_args()

    log("fetch_osm_report.py run started")

    target_dir = args.dry_run if args.dry_run else DOWNLOADS_DIR
    dry_run_label = "  [DRY RUN -- not the real Downloads folder]" if args.dry_run else ""
    if args.dry_run:
        log(f"DRY RUN mode: saving to {target_dir} instead of {DOWNLOADS_DIR}")

    outlook, mapi, inbox = connect_to_outlook()
    osm_folder = find_osm_folder(inbox)

    msg, most_recent_from_sender = find_todays_report(osm_folder)

    if msg is None:
        if most_recent_from_sender is not None:
            try:
                rt = most_recent_from_sender.ReceivedTime
                subj = most_recent_from_sender.Subject
                log(f"Most recent email from {OSM_SENDER_EMAIL}: "
                    f"received {rt}, subject {subj!r}")
            except Exception:
                pass
            sys.exit(
                f"ERROR: No email from {OSM_SENDER_EMAIL} with subject "
                f"{OSM_SUBJECT!r} received TODAY ({datetime.now().date()}). "
                f"The dashboard was NOT refreshed automatically this morning. "
                f"If the report is just running late, re-run this script once "
                f"it arrives, or run the manual Update HRIS Dashboard.bat path."
            )
        else:
            sys.exit(
                f"ERROR: No email at all from {OSM_SENDER_EMAIL} found in "
                f"Inbox/Reports/OSM (checked the 30 most recent items). "
                f"Sender address may have changed, or the mail rule filing "
                f"mail into this folder may have stopped working — check "
                f"Outlook directly."
            )

    log(f"Today's OSM report found: received {msg.ReceivedTime}, subject {msg.Subject!r}")
    save_attachment(msg, target_dir, dry_run_label)

    log("fetch_osm_report.py completed successfully.")


if __name__ == "__main__":
    main()
