"""
fetch_osm_report.py — automated pull of the daily OSM "All Open Tasks by
Team" report attachment from Outlook, for the HRIS dashboard's morning
auto-refresh.

WHAT THIS DOES (and does NOT do):
  - Connects to Outlook via COM, looks first in the mail-rule-filtered
    folder Inbox > Reports > OSM, then falls back to the top-level Inbox
    (added 2026-09-08 — the client-side rule that files the report into
    Reports/OSM only runs while classic Outlook is open, and can lag past
    this script's own morning run, leaving today's report sitting unfiled
    in the top-level Inbox), for today's report email from
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


def _received_local_ymd(msg):
    """(year, month, day) an item was received, in LOCAL time.

    msg.ReceivedTime comes back from COM as a timezone-aware value
    (observed live: UTC, e.g. '...+00:00'). The old code took .year/.month/
    .day straight off that UTC value and compared it to a local date, which
    is only safe well away from midnight. Convert to local time first so a
    report that lands near a day boundary is still bucketed onto the right
    calendar day around the BST/UTC offset.
    """
    rt = msg.ReceivedTime
    try:
        if getattr(rt, "tzinfo", None) is not None:
            rt = rt.astimezone()  # -> machine local time
    except Exception:
        pass
    return rt.year, rt.month, rt.day


def scan_folder_for_today(folder, today, label):
    """Scan the newest items of `folder` for today's OSM report.

    Returns (todays_msg_or_None, most_recent_from_sender_or_None). The
    second value is diagnostic — the most recent matching-sender email in
    this folder regardless of date — so a failure can say *why* ("sender
    matched but it's from yesterday" vs "no email from this sender at all").

    Unlike the old find_todays_report, this does NOT give up at the first
    sender+subject match that isn't dated today. A stray item sorting above
    today's (clock skew, a message re-filed with an odd ReceivedTime, sort
    not applied) must not make a report that IS present look absent. Scans
    up to 60 of the newest items.
    """
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)  # descending — newest first
    except pywintypes.com_error as e:
        log(f"  ({label}: could not sort by ReceivedTime ({e}) — scanning in natural order)")

    most_recent_from_sender = None
    checked = 0
    msg = items.GetFirst()
    while msg is not None:
        checked += 1
        if checked > 60:
            break
        current, msg = msg, items.GetNext()
        try:
            if getattr(current, "Class", None) != 43:  # olMail only
                continue
            sender = (getattr(current, "SenderEmailAddress", "") or "").strip().lower()
            if sender != OSM_SENDER_EMAIL.lower():
                continue
            if most_recent_from_sender is None:
                most_recent_from_sender = current
            subject = (getattr(current, "Subject", "") or "").strip()
            if subject != OSM_SUBJECT:
                log(f"  ({label}: sender matched but subject was {subject!r}, expected {OSM_SUBJECT!r} — skipping)")
                continue
            y, m, d = _received_local_ymd(current)
            if (y, m, d) == (today.year, today.month, today.day):
                return current, most_recent_from_sender
        except Exception as e:
            log(f"  ({label}: error reading an item while scanning: {e})")
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
    today = datetime.now().date()

    # Primary location: the mail-rule-filtered folder Inbox/Reports/OSM.
    osm_folder = find_osm_folder(inbox)
    msg, most_recent_from_sender = scan_folder_for_today(osm_folder, today, "Inbox/Reports/OSM")
    checked_locations = "Inbox/Reports/OSM"

    # Fallback: the client-side rule that files the report out of the
    # top-level Inbox into Reports/OSM only runs while classic Outlook is
    # open, and can lag well past this script's morning run — confirmed
    # live 2026-09-08, when today's 08:00 report sat unfiled in the
    # top-level Inbox past the 09:15 retry. If it isn't in the subfolder
    # yet, check the top-level Inbox itself before giving up.
    if msg is None:
        log("Today's report not found in Inbox/Reports/OSM yet — checking the top-level Inbox in case the mail rule hasn't filed it yet...")
        inbox_msg, inbox_most_recent = scan_folder_for_today(inbox, today, "Inbox")
        checked_locations = "Inbox/Reports/OSM and the top-level Inbox"
        if inbox_msg is not None:
            msg = inbox_msg
        if most_recent_from_sender is None:
            most_recent_from_sender = inbox_most_recent

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
                f"{OSM_SUBJECT!r} received TODAY ({today}). "
                f"Checked {checked_locations}. "
                f"The dashboard was NOT refreshed automatically this morning. "
                f"If the report is just running late, re-run this script once "
                f"it arrives, or run the manual Update HRIS Dashboard.bat path."
            )
        else:
            sys.exit(
                f"ERROR: No email at all from {OSM_SENDER_EMAIL} found in "
                f"{checked_locations} (checked the 60 most recent items in each). "
                f"Sender address may have changed, or the mail rule filing "
                f"mail into Reports/OSM may have stopped working — check "
                f"Outlook directly."
            )

    log(f"Today's OSM report found: received {msg.ReceivedTime}, subject {msg.Subject!r}")
    save_attachment(msg, target_dir, dry_run_label)

    log("fetch_osm_report.py completed successfully.")


if __name__ == "__main__":
    main()
