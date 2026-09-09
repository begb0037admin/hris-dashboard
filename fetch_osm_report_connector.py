#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_osm_report_connector.py -- Codex M365 connector replacement for
fetch_osm_report.py's classic-Outlook-COM attachment fetch.

WHY THIS EXISTS
================
fetch_osm_report.py connects to Outlook via win32com COM and depends on a
client-side Outlook rule + classic Outlook being open (see MIGRATION-IMAP.md
for the 8 Sep 2026 incident this caused). Kevin's standing architecture
decision (9 Sep 2026): all new Microsoft 365 access goes through the Codex
M365 connector, not COM, not IMAP. See
agent-commons/operating-model/CODEX_M365_CONNECTOR_METHOD.md for the method,
and begb0037admin/work-inbox's lane_b_call1.py for the reference
implementation this script deliberately reuses rather than reinvents.

WHAT THIS DOES (mirrors fetch_osm_report.py's contract exactly)
=================================================================
  - Via the Microsoft Outlook Email app connector (codex_apps, personal
    ChatGPT account -- Edu has no email connector attached, same constraint
    already accepted for work-inbox's own mail-connector cutover), in
    READ-ONLY mode: search the Inbox for today's OSM report email
    (OSM_SENDER_EMAIL / OSM_SUBJECT), list its attachments, fetch the
    matching attachment (materialises a short-TTL presigned download URL),
    then downloads that URL directly (plain HTTP GET, no connector
    involved) and saves the bytes to disk.
  - Does NOT parse the spreadsheet, does NOT push to GitHub, does NOT touch
    data/tickets.json -- exactly like fetch_osm_report.py, that work stays
    in import_osm_report.py, unmodified. Point the caller at the saved file
    via OSM_IMPORT_FILE or --file (import_osm_report.py already supports
    both), since this script does not assume a Downloads-folder convention
    tied to any one machine.
  - Saves to a fixed filename ("All Open Tasks by Team - auto.<ext>"),
    matching the existing naming convention, so any tooling that already
    greps for that name keeps working.

REUSE, NOT REINVENTION
=======================
This script imports its connector primitives directly from work-inbox's
lane_b_call1.py (same machine, same account, same already-proven code):
run_codex_json (subprocess invocation + retry/backoff/tree-kill), the JSONL
parser, extract_tool_calls, ReContaminationDetected, READ_VERB_RE/
WRITE_VERB_RE (the verb classification the guard is built on), and
FAILOVER_CODEX_HOME (the personal-account CODEX_HOME already logged into
and proven working for mail). It does NOT import guard_recontamination()
itself or LANE_B_NAMESPACES -- those are scoped to work-inbox's own
calendar/teams/mail domains. Instead this script runs its own narrower
danger-scan restricted to exactly {"microsoft_outlook_email"} (this script
has no legitimate reason to ever touch calendar or Teams tools, so it does
not inherit that trust surface) built on the SAME imported verb regexes, so
verb classification can never drift from work-inbox's own definition even
though the namespace allowlist here is deliberately narrower.

KNOWN COUPLING, DISCLOSED NOT HIDDEN: this script has a hard runtime
dependency on work-inbox's lane_b_call1.py existing at WORKINBOX_REPO_PATH
(default: sibling "work-inbox" clone next to wherever hris-dashboard is
cloned) and exporting the names imported below. If work-inbox ever renames
or removes any of them, this script breaks loudly at import time (exit 2)
-- there is no CI across repos to catch that early. See HANDOVER.md.

EXIT CODES (deliberately distinct, see lane_b_call1.py's 9 Sep 2026 exit-code
-collision incident -- a caller MUST be able to tell these apart without
ambiguity):
  0 = attachment saved successfully.
  1 = GUARD HALT -- a write-verb or off-scope connector tool call was
      observed. Treat exactly like work-inbox's mail-guard HALT: do not
      retry blindly, investigate before relying on this identity again.
  2 = usage / environment error (import failure, bad args, any other
      uncaught exception). Never means a write was detected.
  3 = codex run failed / connector unavailable this cycle after retries
      (headless tool-loading flakiness -- known, documented, retry on the
      normal cadence, not a safety event).
  4 = no matching OSM report email found for today (yet). Same semantics as
      fetch_osm_report.py's sys.exit on this condition -- genuinely nothing
      to fetch, not an error; retry later in the morning's cadence.

Usage:
    python fetch_osm_report_connector.py                  # real run, saves to HRIS_OSM_TARGET_DIR
    python fetch_osm_report_connector.py --dry-run DIR     # saves to DIR instead, for verification
    python fetch_osm_report_connector.py --dry-run DIR --selftest-guard   # pure-function guard check, no codex

Requires: pip install requests  (already confirmed present on the laptop, 9 Sep 2026)
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import re
import sys
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Reuse work-inbox's proven connector primitives -- do not reinvent them.
# ---------------------------------------------------------------------------
WORKINBOX_REPO_PATH = os.environ.get(
    "WORKINBOX_REPO_PATH",
    str(Path(__file__).resolve().parent.parent / "work-inbox"),
)


def _log(msg: str) -> None:
    print(f"[{_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


if WORKINBOX_REPO_PATH not in sys.path:
    sys.path.insert(0, WORKINBOX_REPO_PATH)

try:
    from lane_b_call1 import (  # noqa: E402
        run_codex_json,
        extract_tool_calls,
        final_assistant_text,
        ReContaminationDetected,
        READ_VERB_RE,
        WRITE_VERB_RE,
        FAILOVER_CODEX_HOME,
        SAFETY_RULE,
        CALL1_TIMEOUT_S,
        _utcstamp,
    )
except Exception as _imp_e:  # noqa: BLE001 -- must never look like a guard HALT (exit 1)
    print(
        f"[{_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] FATAL: cannot import "
        f"lane_b_call1 from {WORKINBOX_REPO_PATH!r} ({_imp_e}). This script deliberately "
        f"reuses work-inbox's proven connector/guard code rather than duplicating it -- "
        f"see MIGRATION-IMAP.md 'REUSE, NOT REINVENTION'. Set WORKINBOX_REPO_PATH if "
        f"work-inbox is cloned somewhere else on this machine.",
        file=sys.stderr,
    )
    sys.exit(2)

# ---------------------------------------------------------------------------
# Config -- mirrors fetch_osm_report.py's constants exactly where it matters
# ---------------------------------------------------------------------------
OSM_SENDER_EMAIL = "reports-prd-ldz@saasiteu.com"
OSM_SUBJECT = "Report: All Open Tasks by Team"
SAVED_FILENAME_BASE = "All Open Tasks by Team - auto"

DEFAULT_TARGET_DIR = os.environ.get(
    "HRIS_OSM_TARGET_DIR",
    str(Path(__file__).resolve().parent / "data" / "osm_fetch"),
)

# This script's OWN trust surface: microsoft_outlook_email ONLY. Deliberately
# narrower than lane_b_call1.py's LANE_B_NAMESPACES (calendar+teams+email) --
# this script has no legitimate reason to ever call a calendar/Teams tool, so
# it does not inherit that wider allowance. The verb regexes themselves are
# imported, not redefined, so "what counts as a read/write verb" can never
# drift from work-inbox's own single definition.
OSM_NAMESPACES = {"microsoft_outlook_email"}

OSM_RETRIES = max(1, int(os.environ.get("HRIS_OSM_RETRIES", "4")))
OSM_RETRY_BACKOFF_S = [10, 25, 45, 60]
OSM_TIMEOUT_S = int(os.environ.get("HRIS_OSM_TIMEOUT", str(CALL1_TIMEOUT_S)))

_URL_RE = re.compile(r"^https?://", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Prompt -- task-descriptive (imperative tool-naming does not load codex_apps
# tools, per work-inbox's own 1 Sept 2026 probe finding), rigid "raw JSON
# only" instruction (Layer 1 -- dumb fetch, no reasoning over hostile
# content), explicit write-tool ban, SAFETY_RULE appended. Verified end to
# end 8 Sep 2026 against this exact mailbox (search_messages ->
# list_attachments -> fetch_attachment -> curl the presigned URL, sha256
# byte-identical to the COM-fetched file) -- this prompt reproduces that
# proven three-step chain as a single connector session.
# ---------------------------------------------------------------------------
def build_osm_report_prompt(today_iso: str) -> str:
    return (
        "Using the Microsoft Outlook Email app connector, in READ-ONLY mode, do the "
        f"following three steps in order. Step 1: search my Inbox for the newest email "
        f"from {OSM_SENDER_EMAIL} with subject exactly '{OSM_SUBJECT}' received on "
        f"{today_iso} (today). Step 2: if and only if such an email is found, list its "
        f"attachments. Step 3: fetch the attachment whose filename starts with "
        f"'All Open Tasks by Team' and ends in .xls or .xlsx -- this materialises a "
        f"downloadable file reference; report back its download URL and filename exactly "
        f"as returned. Return ONLY the raw connector result for each step you actually "
        f"performed, as JSON, with no summary, no interpretation, and no prose. If no "
        f"matching email is found for today in step 1, say so plainly in one short "
        f"sentence and stop -- do not perform steps 2 or 3, and do not fetch or guess at "
        f"any other email's attachment. "
        "Do not use any other app or tool. Do not send, reply to, forward, move, delete, "
        "mark as read, categorise, flag, or otherwise modify any message or attachment. "
        f"{SAFETY_RULE}"
    )


# ---------------------------------------------------------------------------
# Narrow, OSM-specific danger-scan. Deliberate near-verbatim structural copy
# of lane_b_call1._scan_tool_calls_for_unexpected -- kept local (not imported)
# ONLY so the namespace allowlist here can be OSM_NAMESPACES (strictly
# narrower) rather than work-inbox's own LANE_B_NAMESPACES. The verb
# classification itself (READ_VERB_RE / WRITE_VERB_RE) is imported, never
# redefined, so it cannot silently drift from work-inbox's single definition.
# ---------------------------------------------------------------------------
def scan_tool_calls_for_unexpected(tool_calls: list[dict]) -> tuple[list[str], list[str]]:
    seen: list[str] = []
    unexpected: list[str] = []
    for tc in tool_calls:
        srv, tool = tc["server"], tc["tool"]
        seen.append(f"{srv}::{tool}")
        if srv != "codex_apps":
            unexpected.append(f"{srv}::{tool} (server != codex_apps)")
            continue
        ns, _dot, leaf = tool.partition(".")
        if not _dot:
            unexpected.append(f"{tool} (no namespace)")
            continue
        if ns not in OSM_NAMESPACES:
            unexpected.append(f"{tool} (off-scope connector namespace '{ns}' -- this "
                              f"script only trusts {sorted(OSM_NAMESPACES)})")
            continue
        if WRITE_VERB_RE.match(leaf):
            unexpected.append(f"{tool} (write verb)")
            continue
        if not READ_VERB_RE.match(leaf):
            unexpected.append(f"{tool} (unrecognised verb -- add to lane_b_call1.READ_VERB_RE "
                              f"if it is genuinely a read, matching work-inbox's own convention)")
            continue
    return seen, unexpected


def _find_first_url(obj, _depth=0) -> str:
    """Recursively search a connector result for the first http(s) URL string --
    the fetch_attachment result's exact field name for the presigned download
    URL is not pinned to a single known key (ConnectorFileReference shape not
    fully documented), so this is intentionally schema-tolerant rather than
    hardcoding one field name and breaking silently if it differs."""
    if _depth > 6 or obj is None:
        return ""
    if isinstance(obj, str):
        return obj if _URL_RE.match(obj.strip()) else ""
    if isinstance(obj, dict):
        for v in obj.values():
            u = _find_first_url(v, _depth + 1)
            if u:
                return u
        return ""
    if isinstance(obj, (list, tuple)):
        for v in obj:
            u = _find_first_url(v, _depth + 1)
            if u:
                return u
        return ""
    return ""


def _find_attachment_filename(tool_calls: list[dict]) -> str:
    """Best-effort filename lookup from list_attachments/fetch_attachment
    results -- used only to pick the right file extension; falls back to
    .xlsx if nothing usable is found (parse_report() in import_osm_report.py
    will fail loudly on a genuinely wrong file, this is not a silent risk)."""
    for tc in tool_calls:
        leaf = tc["tool"].split(".")[-1]
        if leaf not in ("list_attachments", "fetch_attachment"):
            continue
        res = tc.get("result") or {}

        def _walk(o, _d=0):
            if _d > 6 or o is None:
                return ""
            if isinstance(o, str):
                low = o.lower()
                if "all open tasks by team" in low and (low.endswith(".xls") or low.endswith(".xlsx")):
                    return o
                return ""
            if isinstance(o, dict):
                for v in o.values():
                    hit = _walk(v, _d + 1)
                    if hit:
                        return hit
                return ""
            if isinstance(o, (list, tuple)):
                for v in o:
                    hit = _walk(v, _d + 1)
                    if hit:
                        return hit
                return ""
            return ""

        hit = _walk(res)
        if hit:
            return hit
    return ""


class OsmEmailNotFoundToday(RuntimeError):
    """Raised when the connector session ran cleanly (no guard issue) but
    found no matching email for today -- maps to exit 4, distinct from a
    codex/connector technical failure (exit 3) and from a guard HALT (exit 1).
    Retrying THIS SAME invocation would not help; the normal 08:45/09:15/09:45
    scheduled-task cadence is the correct retry mechanism, not this script."""


def fetch_osm_attachment_via_connector(today_iso: str) -> tuple[bytes, str]:
    """Returns (attachment_bytes, filename). Raises ReContaminationDetected on
    a guard HALT, OsmEmailNotFoundToday if genuinely absent, or RuntimeError
    after retries are exhausted (connector unavailable this cycle)."""
    prompt = build_osm_report_prompt(today_iso)
    last_err: Exception | None = None

    for attempt in range(1, OSM_RETRIES + 1):
        _log(f"connector attempt {attempt}/{OSM_RETRIES} (CODEX_HOME={FAILOVER_CODEX_HOME}, "
             f"personal-account-only, same identity already proven for work-inbox mail)")
        try:
            objs, raw = run_codex_json(
                prompt, timeout_s=OSM_TIMEOUT_S, tag="hris_osm",
                codex_home=FAILOVER_CODEX_HOME, max_attempts=2,
            )
        except ReContaminationDetected:
            raise
        except Exception as e:  # noqa: BLE001 -- codex run failed outright this attempt
            last_err = e
            _log(f"attempt {attempt} failed ({e})")
            if attempt < OSM_RETRIES:
                wait = OSM_RETRY_BACKOFF_S[min(attempt - 1, len(OSM_RETRY_BACKOFF_S) - 1)]
                _log(f"backing off {wait}s before retrying")
                time.sleep(wait)
            continue

        # Persist the raw transcript for auditability -- same convention as
        # lane_b_call1.py's own <ts>_call1_<domain>_<identity>_a<n>.jsonl files.
        try:
            debug_dir = Path(__file__).resolve().parent / "data" / "osm_fetch" / "transcripts"
            debug_dir.mkdir(parents=True, exist_ok=True)
            (debug_dir / f"{_utcstamp()}_hris_osm_a{attempt}.jsonl").write_text(raw, encoding="utf-8")
        except OSError:
            pass

        tool_calls = extract_tool_calls(objs)
        seen, unexpected = scan_tool_calls_for_unexpected(tool_calls)
        if unexpected:
            raise ReContaminationDetected(
                f"[hris_osm] unexpected tool call(s): {sorted(set(unexpected))} (seen: {sorted(set(seen))})"
            )

        did_search = any(tc["tool"].split(".")[-1] == "search_messages" for tc in tool_calls)
        did_fetch = any(tc["tool"].split(".")[-1] == "fetch_attachment" for tc in tool_calls)

        if did_fetch:
            for tc in tool_calls:
                if tc["tool"].split(".")[-1] != "fetch_attachment":
                    continue
                url = _find_first_url(tc.get("result"))
                if url:
                    filename = _find_attachment_filename(tool_calls) or "All Open Tasks by Team.xlsx"
                    _log(f"fetch_attachment returned a download URL for {filename!r} -- "
                         f"downloading immediately (presigned URLs are short-TTL)")
                    import requests  # local import: only needed on the success path
                    resp = requests.get(url, timeout=60)
                    resp.raise_for_status()
                    return resp.content, filename
            _log(f"attempt {attempt}: fetch_attachment fired but no URL-shaped field found in its "
                 f"result -- treating as unavailable this attempt (see saved transcript for the "
                 f"real shape; this needs a one-time field-name fix once seen live)")
        elif did_search:
            # search ran, but the model stopped before list_attachments/fetch_attachment --
            # per the prompt, this means step 1 found nothing for today.
            text = final_assistant_text(objs).strip()
            _log(f"attempt {attempt}: search_messages ran but no attachment was fetched -- "
                 f"reading as 'no matching email found for today'. Model's own words: {text[:300]!r}")
            raise OsmEmailNotFoundToday(text or "no matching OSM report email found for today")
        else:
            _log(f"attempt {attempt}: connector tools never fired (availability flakiness, "
                 f"documented and expected some cycles) -- treating as unavailable this attempt")

        if attempt < OSM_RETRIES:
            wait = OSM_RETRY_BACKOFF_S[min(attempt - 1, len(OSM_RETRY_BACKOFF_S) - 1)]
            _log(f"backing off {wait}s before retrying")
            time.sleep(wait)

    raise RuntimeError(
        f"connector produced no usable fetch_attachment result after {OSM_RETRIES} attempts"
        + (f" (last error: {last_err})" if last_err else "")
    )


def cmd_selftest_guard() -> int:
    """Pure-function check of scan_tool_calls_for_unexpected -- no codex, no
    connector, no live risk. Mirrors lane_b_call1.py's own cmd_selftest()
    style for this narrower, OSM-specific guard."""
    fails = []

    def check(name, cond):
        print(("  ok   " if cond else "  FAIL ") + name)
        if not cond:
            fails.append(name)

    def tc(server, tool):
        return {"server": server, "tool": tool, "arguments": {}, "result": None, "error": None, "raw": {}}

    clean = [
        tc("codex_apps", "microsoft_outlook_email.search_messages"),
        tc("codex_apps", "microsoft_outlook_email.list_attachments"),
        tc("codex_apps", "microsoft_outlook_email.fetch_attachment"),
    ]
    _, unexpected = scan_tool_calls_for_unexpected(clean)
    check("clean OSM fetch chain (search/list/fetch) -> no unexpected", not unexpected)

    dirty_send = clean + [tc("codex_apps", "microsoft_outlook_email.send_email")]
    _, unexpected2 = scan_tool_calls_for_unexpected(dirty_send)
    check("send_email present -> flagged unexpected", any("send_email" in u for u in unexpected2))

    dirty_move = clean + [tc("codex_apps", "microsoft_outlook_email.move_message")]
    _, unexpected3 = scan_tool_calls_for_unexpected(dirty_move)
    check("move_message present -> flagged unexpected", any("move_message" in u for u in unexpected3))

    off_scope = clean + [tc("codex_apps", "microsoft_outlook_calendar.list_events")]
    _, unexpected4 = scan_tool_calls_for_unexpected(off_scope)
    check("calendar tool present -> flagged off-scope (narrower than work-inbox's own LANE_B_NAMESPACES)",
          any("off-scope" in u for u in unexpected4))

    off_server = clean + [tc("github", "search_issues")]
    _, unexpected5 = scan_tool_calls_for_unexpected(off_server)
    check("non-codex_apps server -> flagged unexpected", any("server != codex_apps" in u for u in unexpected5))

    print("")
    if fails:
        print(f"RESULT: {len(fails)} FAILED")
        return 1
    print("RESULT: all passed")
    return 0


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description=(
        "Connector-based OSM report attachment fetch (replaces classic-Outlook-COM "
        "fetch_osm_report.py -- see MIGRATION-IMAP.md)"
    ))
    ap.add_argument("--dry-run", metavar="DIR", default=None,
                    help="Save the attachment into DIR instead of HRIS_OSM_TARGET_DIR, "
                         "for verification without touching the live pipeline's input.")
    ap.add_argument("--selftest-guard", action="store_true",
                    help="pure-function checks of the OSM-scoped danger-scan, no codex, no connector")
    args = ap.parse_args(argv)

    if args.selftest_guard:
        return cmd_selftest_guard()

    _log("fetch_osm_report_connector.py run started")
    target_dir = args.dry_run if args.dry_run else DEFAULT_TARGET_DIR
    if args.dry_run:
        _log(f"DRY RUN mode: saving to {target_dir} instead of {DEFAULT_TARGET_DIR}")

    today_iso = _dt.date.today().strftime("%Y-%m-%d")

    try:
        content, filename = fetch_osm_attachment_via_connector(today_iso)
    except OsmEmailNotFoundToday as e:
        _log(f"No matching OSM report email found for today ({today_iso}): {e}")
        _log("The dashboard was NOT refreshed this cycle. If the report is just running "
             "late, this script will be re-run automatically on the normal 08:45/09:15/09:45 "
             "cadence, or run it again once the email arrives.")
        return 4
    except RuntimeError as e:
        _log(f"Connector fetch failed after retries (this cycle only, not a safety event): {e}")
        return 3

    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".xls", ".xlsx"):
        _log(f"WARNING: fetched attachment {filename!r} does not have a recognised .xls/.xlsx "
             f"extension -- saving with .xlsx as a best guess. If this is wrong, "
             f"import_osm_report.py will fail loudly on it rather than silently misreading it.")
        ext = ".xlsx"

    os.makedirs(target_dir, exist_ok=True)
    target_path = os.path.join(target_dir, f"{SAVED_FILENAME_BASE}{ext}")
    with open(target_path, "wb") as fh:
        fh.write(content)
    _log(f"Saved attachment {filename!r} ({len(content)} bytes) -> {target_path}")
    _log("fetch_osm_report_connector.py completed successfully.")
    print(target_path)  # last stdout line = the saved path, for a caller to capture cleanly
    return 0


if __name__ == "__main__":
    # Top-level exception guard -- same discipline as lane_b_call1.py's own
    # 9 Sep 2026 fix, for the exact same reason: an uncaught exception must
    # NEVER exit with the same code (1) this script uses for a genuine guard
    # HALT, or a caller cannot tell "a write was observed" apart from "codex
    # wasn't on PATH in this session" and may react as if mailbox safety were
    # at risk when it wasn't.
    try:
        sys.exit(main(sys.argv[1:]))
    except ReContaminationDetected as _e:
        _log(f"RE-CONTAMINATION GUARD HALT: {_e}")
        sys.exit(1)
    except SystemExit:
        raise
    except Exception as _e:  # noqa: BLE001 -- deliberate catch-all, see comment above
        _log(f"UNCAUGHT EXCEPTION ({type(_e).__name__}): {_e} -- environment/plumbing failure, "
             f"NOT a detected write. Exiting 2, not 1.")
        import traceback
        traceback.print_exc()
        sys.exit(2)
