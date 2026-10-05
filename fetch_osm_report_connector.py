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
  5 = MODEL POLICY VIOLATION (added 10 Sep 2026, Priority 4) -- the model/
      effort selection sourced from codex_model_policy.py (constitution/
      MODEL_POLICY.md) refused to build the codex exec call, e.g. an
      xhigh/max ceiling breach or an unrecognised effort override. This is a
      deterministic code/config bug, NOT connector flakiness -- deliberately
      distinct from exit 3 so a caller does not retry it on the normal
      cadence expecting it to self-resolve; investigate the codex_model_policy
      usage in this script or in work-inbox/lane_b_call1.py instead.

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
    import codex_model_policy  # noqa: E402 -- same sys.path, sibling module to lane_b_call1
    from lane_b_call1 import (  # noqa: E402
        run_codex_json,
        extract_tool_calls,
        final_assistant_text,
        ReContaminationDetected,
        READ_VERB_RE,
        WRITE_VERB_RE,
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

# Added 5 Oct 2026 (Drew) -- root-cause fix for an 11-day (24 Sep - 5 Oct)
# 100%-failure streak, root-caused via a live reproduction on this machine
# (every attempt: "connector tools never fired", not an auth/reauth error).
# work-inbox's 24-25 Sep 2026 "three-identity connector failover ring" work
# (lane_b_call1.py commits 75eea0ba/2675afa7) connected the SAME FAILOVER
# identity (C:\WorkInboxAI\codex-laneb, kevin@lelitte.co.uk) to MULTIPLE M365
# accounts (an Oxford mailbox alongside the personal one). Once the connector
# account-picker became model-mediated, an unqualified "my Inbox" prompt
# stopped reliably producing a tool call at all -- the model has no
# deterministic way to choose an account, so some runs silently issue zero
# tool calls. work-inbox's own fix (`_prompt_for_identity()`) is to name the
# target mailbox explicitly as a single short leading sentence. This script
# was never updated to match when that fix shipped -- it's a narrower,
# single-identity caller that doesn't go through work-inbox's ring/
# _prompt_for_identity() machinery, so it needs the same sentence inlined
# here. Account value matches lane_b_identities.json's "personal-uk" entry
# (same CODEX_HOME this script already hardcodes via FAILOVER_CODEX_HOME).
OSM_M365_ACCOUNT = "kevin.lelitte@admin.ox.ac.uk"

# 5 Oct 2026: pinned explicitly instead of importing lane_b_call1.FAILOVER_CODEX_HOME
# (ring slot 1 of work-inbox's lane_b_identities.json, which became begb0037@ox.ac.uk
# Edu -- Edu cannot run the Outlook connector headless: two 180s timeouts, no tool
# calls). codex-lanec = kevin@lelitte.com (Plus), live-verified working 5 Oct 2026.
_ENV_HOME = os.environ.get("HRIS_CODEX_HOME", "").strip()
# Kevin's policy (1 Oct 2026): begb0037@ox.ac.uk DEFAULT -> kevin@lelitte.co.uk ->
# kevin@lelitte.com. Homes without an auth.json are skipped. HRIS_CODEX_HOME, if
# set, overrides the ring with that single home. MODELS (Kevin, 5 Oct 2026), always
# with -c model_reasoning_effort=high and no fallback substitution:
#   begb (Edu): gpt-6-luna -- the slug the Codex desktop app shows ("GPT-6 Luna High").
#     Headless codex-cli 0.151.0 (npm) rejects it ("Model metadata not found" / 400),
#     but the app-bundled codex.exe (0.159.2) accepts it -- live-verified 5 Oct 2026.
#     So begb runs through the app's bundled binary (newest
#     %LOCALAPPDATA%\OpenAI\Codex\bin\<hash>\codex.exe; override HRIS_CODEX_APP_BIN).
#   lelitte.* (Plus): gpt-5.6-luna via the normal `codex` CLI (policy default).
# A 400 "model not supported" advances the ring immediately (no retry, no wait).
ACCOUNT_RING = [
    {"label": "begb0037@ox.ac.uk", "home": r"C:\WorkInboxAI\codex-laneb", "timeout": 420, "retries": 1,
     "model": "gpt-6-luna", "bin": "app"},
    {"label": "kevin@lelitte.co.uk", "home": os.environ.get("HRIS_CODEX_HOME_LELITTE_CO_UK", r"C:\WorkInboxAI\codex-lelitte-couk"),
     "timeout": 420, "retries": 1, "model": "", "bin": None},
    {"label": "kevin@lelitte.com", "home": r"C:\WorkInboxAI\codex-lanec", "timeout": None, "retries": None,
     "model": "", "bin": None},
]
_ENV_HOME = os.environ.get("HRIS_CODEX_HOME", "").strip()
if _ENV_HOME:
    ACCOUNT_RING = [{"label": "HRIS_CODEX_HOME", "home": _ENV_HOME, "timeout": None, "retries": None,
                     "model": "", "bin": None}]


def _app_codex_bin():
    import glob
    ov = os.environ.get("HRIS_CODEX_APP_BIN", "").strip()
    if ov and os.path.isfile(ov):
        return ov
    base = os.path.join(os.environ.get("LOCALAPPDATA", r"C:\Users\begb0037.AD-OAK\AppData\Local"),
                        "OpenAI", "Codex", "bin", "*", "codex.exe")
    found = sorted(glob.glob(base), key=os.path.getmtime, reverse=True)
    return found[0] if found else None


_URL_RE = re.compile(r"^https?://", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Prompt -- task-descriptive (imperative tool-naming does not load codex_apps
# tools, per work-inbox's own 1 Sept 2026 probe finding), rigid "raw JSON
# only" instruction (Layer 1 -- dumb fetch, no reasoning over hostile
# content), explicit write-tool ban, SAFETY_RULE appended. Verified end to
# end 8 Sep 2026 against this exact mailbox (search_messages ->
# list_attachments -> fetch_attachment -> curl the presigned URL, sha256
# byte-identical to the COM-fetched file) -- this prompt reproduces that
# proven three-step chain as a single connector session. 5 Oct 2026: leading
# account-targeting sentence added (see OSM_M365_ACCOUNT comment above) --
# wording deliberately mirrors work-inbox's own `_prompt_for_identity()`
# verbatim ("Use only the Oxford Microsoft 365 mailbox {account}, not
# Personal.") since that exact wording is the one already proven live to
# resolve the account-picker ambiguity in work-inbox's own ring.
# ---------------------------------------------------------------------------
def build_osm_report_prompt(today_iso: str) -> str:
    return (
        f"Use only the Oxford Microsoft 365 mailbox {OSM_M365_ACCOUNT}, not Personal. "
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


class ModelUnsupported(RuntimeError):
    """400 model-not-supported: advance the ring immediately, never retry."""


class OsmEmailNotFoundToday(RuntimeError):
    """Raised when the connector session ran cleanly (no guard issue) but
    found no matching email for today -- maps to exit 4, distinct from a
    codex/connector technical failure (exit 3) and from a guard HALT (exit 1).
    Retrying THIS SAME invocation would not help; the normal 08:45/09:15/09:45
    scheduled-task cadence is the correct retry mechanism, not this script."""


def fetch_osm_attachment_via_connector(today_iso: str) -> tuple[bytes, str]:
    """Walk ACCOUNT_RING; guard HALT / policy violation / email-not-found are
    terminal (propagate); a RuntimeError (auth/timeout/no tools) advances."""
    import lane_b_call1 as _lb
    _ORIG_CODEX_BIN = _lb.CODEX_BIN
    errs = []
    notfound = None
    for acct in ACCOUNT_RING:
        home = acct["home"]
        if not (Path(home) / "auth.json").is_file():
            _log(f"[ring] skipping {acct['label']}: no auth.json in {home} (needs `codex login` into that CODEX_HOME)")
            errs.append(f"{acct['label']}: skipped")
            continue
        _lb.CODEX_MODEL = acct.get("model") or ""  # "" = codex_model_policy default (gpt-5.6-luna, effort high)
        _lb.CODEX_BIN = (_app_codex_bin() or _ORIG_CODEX_BIN) if acct.get("bin") == "app" else _ORIG_CODEX_BIN
        _log(f"[ring] trying {acct['label']} (CODEX_HOME={home}, model={_lb.CODEX_MODEL or 'gpt-5.6-luna'}/high bin={_lb.CODEX_BIN})")
        try:
            return _fetch_from_home(today_iso, home, acct["timeout"] or OSM_TIMEOUT_S, acct["retries"] or OSM_RETRIES)
        except OsmEmailNotFoundToday as e:
            # A "not found" from one account is NOT authoritative (gpt-6-luna on begb
            # returned a false negative 5 Oct 2026) -- try the next account; only
            # raise it if no account fetched anything.
            _log(f"[ring] {acct['label']} reported no email for today ({str(e)[:80]!r}) -- advancing to confirm")
            notfound = e
            errs.append(f"{acct['label']}: not found")
        except RuntimeError as e:
            if isinstance(e, ReContaminationDetected):
                raise
            _log(f"[ring] {acct['label']} failed ({e}) -- advancing")
            errs.append(f"{acct['label']}: {e}")
    if notfound is not None:
        raise notfound
    raise RuntimeError("all accounts in the ring failed: " + "; ".join(errs))


def _fetch_from_home(today_iso: str, home: str, timeout_s: int, retries: int) -> tuple[bytes, str]:
    """Returns (attachment_bytes, filename). Raises ReContaminationDetected on
    a guard HALT, OsmEmailNotFoundToday if genuinely absent, or RuntimeError
    after retries are exhausted (connector unavailable this cycle)."""
    prompt = build_osm_report_prompt(today_iso)
    last_err: Exception | None = None

    for attempt in range(1, retries + 1):
        _log(f"connector attempt {attempt}/{retries} (CODEX_HOME={home}, "
             f"personal-account-only, same identity already proven for work-inbox mail)")
        try:
            # workload_class="high": MODEL_POLICY.md precedence rule -- this
            # goes through the microsoft_outlook_email connector namespace,
            # which is write-capable (guarded by OSM_NAMESPACES/the verb
            # regexes, not excluded from the tool surface), so it classifies
            # High even though this specific prompt is read-only. Mirrors
            # work-inbox/lane_b_call1.py's own call sites -- same policy,
            # same reasoning, sourced from the same shared module.
            objs, raw = run_codex_json(
                prompt, timeout_s=timeout_s, tag="hris_osm",
                codex_home=home, max_attempts=2,
                workload_class="high",
            )
        except ReContaminationDetected:
            raise
        except codex_model_policy.ModelPolicyViolation:
            # MUST NOT be caught by the generic `except Exception` below --
            # touchpoint-1 Codex review finding, 10 Sep 2026 (same reasoning
            # as work-inbox/lane_b_call1.py's _fetch_domain_one_identity()):
            # a policy violation is a deterministic code/config bug, not
            # connector flakiness. Retrying it retries times wastes the
            # whole retry budget and would exhaust into the generic
            # RuntimeError -> exit 3 ("not a safety event") path in main(),
            # completely mischaracterising a real misconfiguration as
            # ordinary headless-connector flakiness. Re-raise uncaught so
            # main() can map it to its own distinct exit code (5) instead.
            raise
        except Exception as e:  # noqa: BLE001 -- codex run failed outright this attempt
            last_err = e
            _log(f"attempt {attempt} failed ({e})")
            if attempt < retries:
                wait = OSM_RETRY_BACKOFF_S[min(attempt - 1, len(OSM_RETRY_BACKOFF_S) - 1)]
                _log(f"backing off {wait}s before retrying")
                time.sleep(wait)
            continue

        _errtxt = " ".join(str(o) for o in objs if isinstance(o, dict)
                            and str(o.get("type", "")).lower() in ("error", "turn.failed")).lower()
        if "not supported" in _errtxt and "model" in _errtxt:
            raise ModelUnsupported(f"model not supported on {home} (HTTP 400) -- failing fast, advancing ring")
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

        did_search = any(tc["tool"].split(".")[-1] in ("search_messages", "list_messages") for tc in tool_calls)
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

        if attempt < retries:
            wait = OSM_RETRY_BACKOFF_S[min(attempt - 1, len(OSM_RETRY_BACKOFF_S) - 1)]
            _log(f"backing off {wait}s before retrying")
            time.sleep(wait)

    raise RuntimeError(
        f"connector produced no usable fetch_attachment result after {retries} attempts"
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
    except ReContaminationDetected as e:
        # ReContaminationDetected is a RuntimeError subclass, so this must come
        # before the generic RuntimeError branch below. A genuine write/off-scope
        # tool call is a guard HALT (exit 1), not connector flakiness (exit 3),
        # and the scheduled-task wrapper must disable the task rather than retry
        # blindly on the normal cadence.
        _log(f"GUARD HALT -- unexpected/write connector tool call observed: {e}")
        return 1
    except codex_model_policy.ModelPolicyViolation as e:
        # MUST be caught here, before the generic `except RuntimeError` below,
        # or it silently maps to exit 3 ("not a safety event") -- wrong, this
        # is a real code/config bug that will not self-resolve on the normal
        # retry cadence. See exit code 5's own docstring entry above.
        _log(f"MODEL POLICY VIOLATION -- {e}")
        _log("This is a code/config bug, NOT connector flakiness -- do not expect this "
             "to self-resolve on the normal 08:45/09:15/09:45 cadence. Investigate "
             "codex_model_policy.py usage in this script or in work-inbox/lane_b_call1.py.")
        return 5
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
