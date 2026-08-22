"""
check_last_run_status.py -- tells Run_HRIS_Auto_Refresh.bat whether it
should actually run, or skip, on this invocation.

Why this exists (22 Aug 2026): the morning trigger moved from a single
08:30 run to three attempts (08:45 primary, 09:15/09:45 retries) because
the Outlook mail rule that files the OSM report into Inbox/Reports/OSM
doesn't always finish before 08:30 -- confirmed live that morning (see
RESUME.md, "8:30 discrepancy investigation"). Running the full
fetch+update pipeline a second or third time on a morning that ALREADY
succeeded would be wasteful (another live Outlook COM connection, another
GitHub push) and, worse, would overwrite data/last_automated_run.json's
accurate first-success timestamp with a later, redundant one -- exactly
the "blur multiple attempts together confusingly" outcome Kevin asked to
avoid. So every retry trigger checks this FIRST.

Prints exactly one line to stdout:
  "SKIP"                       -- today's run already succeeded; the
                                   caller should do nothing further.
  "RUN:<attempt>:<max>"        -- today has not yet succeeded; caller
                                   should proceed. <attempt> is a
                                   wall-clock-time bucket (1/2/3, matching
                                   the three scheduled trigger times), not
                                   a persisted counter -- avoids any
                                   cross-day state file that could get out
                                   of sync. <max> is always 3, matching
                                   the number of scheduled triggers.

Never exits non-zero and never lets an unrelated problem (network hiccup,
GitHub briefly unreachable, malformed JSON) silently turn into a skipped
morning refresh -- any failure to answer confidently defaults to
"RUN:1:3" so the real pipeline still gets a chance to run. This script
only ever *reads* -- it never writes data/last_automated_run.json itself.

Usage: python check_last_run_status.py
"""
import sys
from datetime import datetime, date, time

try:
    import requests
except ImportError:
    # Consistent with the "never silently skip a real run" rule above --
    # if we can't even check, assume RUN rather than exiting non-zero and
    # tripping up the caller's error handling.
    print("RUN:1:3")
    sys.exit(0)

RAW_URL = "https://raw.githubusercontent.com/begb0037admin/hris-dashboard/main/data/last_automated_run.json"
MAX_ATTEMPTS = 3

# Wall-clock buckets matching the three scheduled trigger times
# (08:45 primary, 09:15 retry 1, 09:45 retry 2). Deliberately time-based
# rather than a persisted per-day counter file -- stateless, so it can't
# drift out of sync from a stray manual test run, a missed trigger, or a
# clock change, and needs no reset-on-new-day logic.
def attempt_bucket(now):
    if now < time(9, 0):
        return 1
    elif now < time(9, 30):
        return 2
    else:
        return 3


def main():
    attempt = attempt_bucket(datetime.now().time())
    try:
        resp = requests.get(f"{RAW_URL}?t={datetime.now().strftime('%Y%m%d%H%M%S')}", timeout=15)
        if resp.status_code != 200:
            print(f"RUN:{attempt}:{MAX_ATTEMPTS}")
            return
        data = resp.json()
        status = data.get("status", "")
        ts = data.get("timestamp", "") or ""
        succeeded_today = status == "success" and ts[:10] == date.today().isoformat()
        print("SKIP" if succeeded_today else f"RUN:{attempt}:{MAX_ATTEMPTS}")
    except Exception:
        print(f"RUN:{attempt}:{MAX_ATTEMPTS}")


if __name__ == "__main__":
    main()
