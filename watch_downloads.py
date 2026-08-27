"""
watch_downloads.py -- HRIS Dashboard "just save the export" watcher (Option 3)

Polls the Downloads folder every ~60s. When a genuinely new
"All Open Tasks by Team*.xlsx / *.xls" file appears and its size has
settled, it runs the SAME import_osm_report.py the manual .bat and the
morning Scheduled Task run, which pushes data/tickets.json to GitHub.
Kevin's whole manual mid-day step becomes: save the export to Downloads.

Design / safety (see hris-dashboard RESUME.md, 27 Aug 2026 session):

  * NEVER touches the morning automation. The Scheduled Task
    "HRIS Dashboard Morning Refresh" saves its file as
    "All Open Tasks by Team - auto.xls"; this watcher hard-excludes any
    name containing "- auto", and also stays completely idle during the
    configurable weekday quiet-hours window (default 08:40-10:00) so it
    can never run an import at the same moment as that task. Files that
    appear during quiet hours are held and imported once the window ends.

  * Dedupe. A state file records {name, size, sha256} of every file it
    has already imported (last 30). It never imports the same bytes
    twice and never double-fires on the morning "- auto" file even if
    the exclusion above were removed.

  * Settle detection. A file must be size-stable across `settle_polls`
    consecutive polls (default 2 => ~2 min) before it is eligible, so a
    still-downloading file is not read mid-write.

  * Runs the hardened import_osm_report.py, which adds: lazy xlrd,
    locked/partial-file detection + retry, an advisory single-writer
    lock shared by all three update paths, and a push race-retry on the
    GitHub Contents API. A watcher-triggered import reaches main exactly
    the same way the morning task's does -- directly, auto, one
    tickets.json commit -- and is reverted the same way: `git revert
    <sha>` or re-PUT the previous blob; Pages rebuilds in ~60-90s.
    index.html is never touched by this path.

  * Kill switch: set "enabled": false in watch_downloads.config.json
    (watcher exits within one poll), or delete the Startup .vbs and
    taskkill the pythonw process.

Stdlib only. Launch hidden via "Start HRIS Downloads Watcher.vbs"
(pythonw, windowStyle 0) from the Startup folder.
Every log line is timestamped (estate-wide requirement).
"""

import os
import sys
import json
import time
import glob
import hashlib
import subprocess
import urllib.request
from datetime import datetime, time as dtime

HERE       = os.path.dirname(os.path.abspath(__file__))
CONFIG     = os.path.join(HERE, "watch_downloads.config.json")
STATE_DIR  = os.path.join(os.environ.get("LOCALAPPDATA", HERE), "hris-downloads-watcher")
STATE_FILE = os.path.join(STATE_DIR, "state.json")
LOG_FILE   = os.path.join(STATE_DIR, "watcher.log")
PID_FILE   = os.path.join(STATE_DIR, "watcher.pid")
PINNED_PY  = os.path.join(STATE_DIR, "import_osm_report.pinned.py")
LOG_MAX    = 1_000_000

DEFAULTS = {
    "enabled": True,
    "script_sha": "",                       # commit SHA of import_osm_report.py to run (pinned)
    "poll_seconds": 60,
    "settle_polls": 2,
    "downloads_dir": r"C:\Users\admin\Downloads",
    "patterns": ["All Open Tasks by Team*.xlsx", "All Open Tasks by Team*.xls"],
    "exclude_substrings": ["- auto", "~$", ".crdownload", ".tmp", ".part"],
    "quiet_hours_weekday": ["08:40", "10:00"],
    "state_keep": 30,
    "ensure_deps": True,
}

# ── logging ─────────────────────────────────────────────────────────────────

def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        os.makedirs(STATE_DIR, exist_ok=True)
        if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > LOG_MAX:
            bak = LOG_FILE + ".1"
            if os.path.exists(bak):
                os.remove(bak)
            os.replace(LOG_FILE, bak)
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass
    try:
        print(line, flush=True)
    except Exception:
        pass

# ── config / state ─────────────────────────────────────────────────────────

def load_config():
    cfg = dict(DEFAULTS)
    try:
        with open(CONFIG, encoding="utf-8") as fh:
            cfg.update(json.load(fh))
    except FileNotFoundError:
        log(f"WARNING: {CONFIG} not found -- using built-in defaults")
    except (OSError, json.JSONDecodeError) as e:
        log(f"WARNING: could not read config ({e}) -- using built-in defaults")
    return cfg

def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return {"imported": []}

def save_state(state, keep):
    state["imported"] = state["imported"][-keep:]
    try:
        os.makedirs(STATE_DIR, exist_ok=True)
        tmp = STATE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(state, fh, indent=2)
        os.replace(tmp, STATE_FILE)
    except OSError as e:
        log(f"WARNING: could not save state: {e}")

# ── single instance ────────────────────────────────────────────────────────

def _pid_alive(pid):
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
            capture_output=True, text=True, timeout=15,
        ).stdout
        return str(pid) in out
    except Exception:
        return False

def claim_single_instance():
    try:
        if os.path.exists(PID_FILE):
            old = open(PID_FILE, encoding="utf-8").read().strip()
            if old.isdigit() and int(old) != os.getpid() and _pid_alive(int(old)):
                log(f"another watcher (pid {old}) is already running -- exiting")
                return False
        os.makedirs(STATE_DIR, exist_ok=True)
        with open(PID_FILE, "w", encoding="utf-8") as fh:
            fh.write(str(os.getpid()))
        return True
    except OSError as e:
        log(f"WARNING: could not claim pid file ({e}) -- continuing anyway")
        return True

# ── helpers ────────────────────────────────────────────────────────────────

def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

def in_quiet_hours(cfg, now=None):
    now = now or datetime.now()
    if now.weekday() > 4:  # Sat/Sun
        return False
    try:
        a, b = cfg["quiet_hours_weekday"]
        ah, am = (int(x) for x in a.split(":"))
        bh, bm = (int(x) for x in b.split(":"))
        return dtime(ah, am) <= now.time() <= dtime(bh, bm)
    except Exception:
        return False

def excluded(name, cfg):
    low = name.lower()
    return any(s.lower() in low for s in cfg["exclude_substrings"])

def candidates(cfg):
    found = []
    for pat in cfg["patterns"]:
        found.extend(glob.glob(os.path.join(cfg["downloads_dir"], pat)))
    out = []
    for p in found:
        n = os.path.basename(p)
        if excluded(n, cfg):
            continue
        try:
            st = os.stat(p)
        except OSError:
            continue
        out.append((p, st.st_size, st.st_mtime))
    return out

def ensure_pinned_script(cfg):
    """Download import_osm_report.py at the pinned SHA. Returns local path or None."""
    sha = (cfg.get("script_sha") or "").strip()
    if not sha:
        log("ERROR: config.script_sha is empty -- cannot run imports until it is set")
        return None
    marker = PINNED_PY + ".sha"
    have = ""
    if os.path.exists(PINNED_PY) and os.path.exists(marker):
        have = open(marker, encoding="utf-8").read().strip()
    if have == sha and os.path.getsize(PINNED_PY) > 500:
        return PINNED_PY
    url = (f"https://raw.githubusercontent.com/begb0037admin/hris-dashboard/"
           f"{sha}/import_osm_report.py?cb={int(time.time())}")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "hris-downloads-watcher"})
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read().decode("utf-8")
    except Exception as e:
        log(f"ERROR: could not download import_osm_report.py @ {sha[:12]}: {e}")
        return None
    if len(body) < 500 or "def main()" not in body:
        log(f"ERROR: downloaded import_osm_report.py @ {sha[:12]} looks wrong "
            f"({len(body)} bytes) -- refusing to use it")
        return None
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(PINNED_PY, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(body)
    with open(marker, "w", encoding="utf-8") as fh:
        fh.write(sha)
    log(f"fetched import_osm_report.py pinned to {sha[:12]} ({len(body)} bytes)")
    return PINNED_PY

def ensure_deps():
    try:
        subprocess.run([sys.executable, "-c", "import openpyxl, requests"],
                       capture_output=True, timeout=20, check=True)
        return
    except Exception:
        pass
    log("installing missing Python deps (openpyxl requests xlrd)...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet",
                        "--disable-pip-version-check", "openpyxl", "requests", "xlrd"],
                       capture_output=True, text=True, timeout=300)
    except Exception as e:
        log(f"WARNING: dep install failed ({e}) -- imports may fail until deps are present")

def run_import(script_path, xlsx_path):
    env = dict(os.environ)
    env["OSM_IMPORT_FILE"] = xlsx_path
    env["OSM_IMPORT_NONINTERACTIVE"] = "1"
    log(f"running import for: {os.path.basename(xlsx_path)}")
    try:
        proc = subprocess.run(
            [sys.executable, script_path],
            capture_output=True, text=True, timeout=600, env=env,
            stdin=subprocess.DEVNULL,
        )
    except subprocess.TimeoutExpired:
        log("ERROR: import timed out after 600s")
        return False
    for ln in (proc.stdout or "").splitlines():
        log(f"  | {ln}")
    for ln in (proc.stderr or "").splitlines():
        log(f"  ! {ln}")
    ok = proc.returncode == 0
    log(f"import {'OK' if ok else 'FAILED (exit %d)' % proc.returncode}")
    return ok

# ── one poll ───────────────────────────────────────────────────────────────

def scan_once(cfg, state, seen):
    """One polling pass. Mutates `state` and `seen` in place.
    Returns list of (name, ok) for any imports actually run this pass."""
    ran = []
    done_keys = {(e["name"], e["size"], e["sha256"]) for e in state["imported"]}
    live_paths = set()

    for path, size, mtime in candidates(cfg):
        live_paths.add(path)
        prev = seen.get(path)
        if not prev or prev[0] != size or prev[1] != mtime:
            seen[path] = [size, mtime, 0]
            continue
        prev[2] += 1  # unchanged this poll
        if prev[2] < cfg["settle_polls"]:
            continue

        name = os.path.basename(path)
        try:
            digest = sha256_of(path)
        except OSError as e:
            log(f"could not hash {name} yet ({e}) -- will retry")
            seen[path] = [size, mtime, 0]
            continue

        if (name, size, digest) in done_keys:
            continue  # already imported these exact bytes

        if in_quiet_hours(cfg):
            log(f"{name} is ready but it is the morning auto-refresh window "
                f"-- deferring until after {cfg['quiet_hours_weekday'][1]}")
            seen[path][2] = cfg["settle_polls"]  # stay eligible, re-check next poll
            continue

        script = ensure_pinned_script(cfg)
        if not script:
            log(f"{name} ready but no runnable import script -- will retry next poll")
            continue

        ok = run_import(script, path)
        state["imported"].append({
            "name": name, "size": size, "mtime": mtime, "sha256": digest,
            "status": "ok" if ok else "failed",
            "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        save_state(state, cfg["state_keep"])
        done_keys.add((name, size, digest))
        ran.append((name, ok))
        if ok:
            log("dashboard updated -- allow ~60-90s for GitHub Pages to rebuild")

    for gone in [p for p in seen if p not in live_paths]:
        del seen[gone]
    return ran

# ── main loop ──────────────────────────────────────────────────────────────

def seed_state():
    """Record every currently-present matching export as already-imported, so a
    freshly-deployed watcher does NOT re-import files that are already live.
    Run once at deploy time: python watch_downloads.py --seed"""
    cfg = load_config()
    state = load_state()
    done = {(e["name"], e["size"], e["sha256"]) for e in state["imported"]}
    added = 0
    for path, size, mtime in candidates(cfg):
        name = os.path.basename(path)
        try:
            digest = sha256_of(path)
        except OSError:
            continue
        if (name, size, digest) in done:
            continue
        state["imported"].append({
            "name": name, "size": size, "mtime": mtime, "sha256": digest,
            "status": "seeded",
            "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        added += 1
        log(f"seeded (will not re-import): {name}")
    save_state(state, cfg["state_keep"])
    log(f"seed complete -- {added} file(s) recorded as already handled")

def main():
    if "--seed" in sys.argv[1:]:
        seed_state()
        return
    log("=" * 60)
    log(f"watcher starting (pid {os.getpid()}, python {sys.version.split()[0]})")
    if not claim_single_instance():
        return
    cfg = load_config()
    if not cfg.get("enabled", True):
        log("config.enabled is false -- exiting")
        return
    log(f"downloads_dir={cfg['downloads_dir']}  poll={cfg['poll_seconds']}s  "
        f"settle={cfg['settle_polls']}  quiet_weekday={cfg['quiet_hours_weekday']}  "
        f"script_sha={cfg.get('script_sha','')[:12] or '(unset)'}")
    if cfg.get("ensure_deps", True):
        ensure_deps()

    state = load_state()
    seen = {}  # path -> [size, mtime, stable_count]

    while True:
        try:
            cfg = load_config()
            if not cfg.get("enabled", True):
                log("config.enabled turned false -- exiting")
                return
            scan_once(cfg, state, seen)
        except Exception as e:  # never let the loop die
            log(f"loop error (continuing): {type(e).__name__}: {e}")

        time.sleep(max(10, int(cfg.get("poll_seconds", 60))))

if __name__ == "__main__":
    main()
